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

use crate::addr::{self, CellRef, Rect};
use crate::engine::{ConstSource, Engine};
use crate::fastmap::{FastMap, FastSet};
use crate::formula::{self, ParseCtx};
use crate::graph::{self, Csr};
use crate::ir::{
    Binding, BuildInfo, BuildOptions, CellDef, CellKind, Expr, Meta, Program, Queryable,
};
use crate::model::{Model, ModelCell};
use crate::options::{DateSystem, EngineOptions, Iteration, Locale, Search};
use crate::value::Value;
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
    pub fn parse(s: &str) -> Self {
        Self {
            target: s.trim().to_string(),
            name: None,
        }
    }

    pub fn named(name: &str, target: &str) -> Self {
        Self {
            target: target.trim().to_string(),
            name: Some(name.trim().to_string()),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Spec {
    pub inputs: Vec<BindingSpec>,
    /// Empty: one binding per sheet covering its formula cells.
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
    /// Which ranges a caller may ask for besides the outputs.
    pub query: QueryLevel,
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
/// each one reads a formula differently from the reference (LibreOffice 26.8), so the user opts in.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Features {
    /// `_xlfn._LONGTEXT(…)` with constant text arguments: LibreOffice 26.2 wrote a text literal of
    /// more than 255 characters as pieces joined by it. The pieces are joined into one literal.
    pub longtext: bool,
}

impl Features {
    /// Every feature the configuration can name.
    pub const NAMES: &'static [&'static str] = &["LONGTEXT"];

    /// The features of an `enable_features` list (names in any case); only those listed are on.
    pub fn parse(names: &[String]) -> Result<Self, String> {
        let mut f = Self::default();
        for n in names {
            match n.trim().to_ascii_uppercase().as_str() {
                "LONGTEXT" => f.longtext = true,
                other => {
                    return Err(format!(
                        "unknown feature `{other}` in enable_features (known: {})",
                        Self::NAMES.join(", ")
                    ));
                }
            }
        }
        Ok(f)
    }
}

/// How much of the workbook the program keeps for run-time queries.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum QueryLevel {
    /// Only the declared outputs; everything else is pruned.
    #[default]
    Declared,
    /// The declared outputs plus the workbook's defined names.
    Named,
    /// Any range.
    Any,
}

impl QueryLevel {
    pub fn parse(s: &str) -> Option<Self> {
        match s.trim().to_ascii_lowercase().as_str() {
            "declared" => Some(Self::Declared),
            "named" => Some(Self::Named),
            "any" => Some(Self::Any),
            _ => None,
        }
    }
}

impl Default for Spec {
    fn default() -> Self {
        Self {
            inputs: Vec::new(),
            outputs: Vec::new(),
            date_system: None,
            iteration: None,
            locale: None,
            search: None,
            now: None,
            host: Default::default(),
            rand_seed: None,
            fold: true,
            query: QueryLevel::Declared,
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
    /// Warnings of the build (text outside the scripts the collation orders as LibreOffice does…).
    pub warnings: Vec<String>,
    /// Every function called in the workbook's formulas and how many times
    /// (including the ones the engine does not implement).
    pub function_uses: std::collections::BTreeMap<String, usize>,
}

#[derive(Debug)]
pub struct CompileError(pub String);

impl std::fmt::Display for CompileError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for CompileError {}

fn err<T>(msg: impl Into<String>) -> Result<T, CompileError> {
    Err(CompileError(msg.into()))
}

const MAX_BINDING_CELLS: u64 = 10_000_000;

/// Resolve a user-facing target to a rectangle.
pub fn resolve_target(model: &Model, target: &str) -> Result<Rect, CompileError> {
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
fn check_bindings(model: &Model, spec: &Spec) -> Result<(), CompileError> {
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

#[derive(Default)]
struct Flags {
    dynamic: bool,
    volatile: bool,
}

/// The peak resident memory of the process so far, in MB, where `/proc` reports it.
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

/// Phase timing and peak memory on stderr when `SHEETC_TRACE` is set.
fn trace(t0: &std::time::Instant, label: &str) {
    if std::env::var_os("SHEETC_TRACE").is_some() {
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
    /// Functions of LibreOffice the engine lacks, with the first cell that calls each.
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
                    && let Some(id) = crate::functions::lookup(crate::ir::Ns::Core, "_ARRAY")
                {
                    e = Expr::Call(id, vec![e]);
                }
                let mut holds_subtotal = false;
                e.walk(&mut |x| match x {
                    Expr::Unknown(call) => {
                        let name = &call.0;
                        out.unsupported.insert(name.to_string());
                        if crate::functions::in_catalogue(name) {
                            out.missing.entry(name.to_string()).or_insert(at);
                        }
                        *out.function_uses.entry(name.to_string()).or_default() += 1;
                    }
                    Expr::Call(id, args) => {
                        let name = id.def().name;
                        if name == "TEXT"
                            && let Some(Expr::Const(Value::Text(code))) = args.get(1)
                            && crate::format::has_foreign_language_tag(code)
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
                out.syntax.push((at, f.to_string(), e.to_string()));
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

pub fn compile(model: &Model, spec: &Spec) -> Result<(Program, Report), CompileError> {
    let t0 = std::time::Instant::now();
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
    // Lowering a formula reads only the model: the cells go to as many threads as
    // the machine has, and what each one found is merged afterwards (the result is
    // sorted below, so the order they finish in does not matter).
    let workers = std::thread::available_parallelism()
        .map_or(1, usize::from)
        .min(8);
    let chunk = formula_cells.len().div_ceil(workers).max(1);
    let lowered: Vec<Lowered> = std::thread::scope(|scope| {
        let handles: Vec<_> = formula_cells
            .chunks(chunk)
            .map(|part| scope.spawn(move || lower_chunk(model, part, calc_syntax, spec.features)))
            .collect();
        handles
            .into_iter()
            .map(|h| h.join().expect("a lowering thread panicked"))
            .collect()
    });
    let mut missing: std::collections::BTreeMap<String, CellRef> = Default::default();
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

    trace(&t0, "parsed formulas");
    if !model.unsupported_options.is_empty() && !spec.inspect {
        return Err(CompileError(format!(
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
                format!("{sheet}!{} `{f}` ({why})", crate::addr::a1(at.row, at.col))
            })
            .collect();
        return Err(CompileError(format!(
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
                format!("{sheet}!{}: {what}", crate::addr::a1(at.row, at.col))
            })
            .collect();
        return Err(CompileError(format!(
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
                    crate::addr::a1(at.row, at.col)
                )
            })
            .collect();
        return Err(CompileError(format!(
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
    if !input_nodes.is_empty() {
        // An input replaces whatever formula or constant sat in its cell.
        let set: FastSet<CellRef> = input_nodes.iter().map(|(a, _)| *a).collect();
        formulas.retain(|(a, _)| !set.contains(a));
        cached.retain(|a, _| !set.contains(a));
        formulas.extend(input_nodes);
        formulas.sort_unstable_by_key(|(at, _)| *at);
        formulas.dedup_by_key(|(at, _)| *at);
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
    if spec.outputs.is_empty() && spec.query == QueryLevel::Declared {
        let mut bbs: Vec<Option<(u32, u32, u32, u32)>> = vec![None; model.sheets.len()];
        for (at, k) in &formulas {
            if matches!(k, CellKind::Formula(_)) && Some(at.sheet) != spec.skip_sheet {
                let bb = &mut bbs[at.sheet as usize];
                *bb = Some(match *bb {
                    None => (at.row, at.col, at.row, at.col),
                    Some((r0, c0, r1, c1)) => (
                        r0.min(at.row),
                        c0.min(at.col),
                        r1.max(at.row),
                        c1.max(at.col),
                    ),
                });
            }
        }
        for (s, bb) in bbs.into_iter().enumerate() {
            if let Some((r0, c0, r1, c1)) = bb {
                let rect = Rect::new(s as u16, r0, c0, r1, c1);
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

    // Ranges kept for run-time queries.
    let mut query_rects: Vec<Rect> = Vec::new();
    let mut named: Vec<Binding> = Vec::new();
    // Defined names that can be read by name: all but the plain-range ones are
    // reported, and none may point into the settings sheet.
    if spec.query != QueryLevel::Declared {
        let mut skipped = Vec::new();
        for name in model.names.keys() {
            match resolve_target(model, name) {
                Ok(rect) if Some(rect.sheet) == spec.skip_sheet => {}
                Ok(rect) if rect.area() <= MAX_BINDING_CELLS => named.push(Binding {
                    name: name.clone(),
                    rect,
                }),
                _ => skipped.push(name.as_str()),
            }
        }
        named.sort_by(|a, b| a.name.cmp(&b.name));
        if !skipped.is_empty() {
            skipped.sort_unstable();
            diags.push(format!(
                "defined names that are not a plain range cannot be read by name: {}",
                skipped.join(", ")
            ));
        }
    }
    match spec.query {
        QueryLevel::Declared => {}
        QueryLevel::Named => query_rects.extend(named.iter().map(|b| b.rect)),
        QueryLevel::Any => {
            let mut ext = vec![(0u32, 0u32); model.sheets.len().max(1)];
            for at in model.cells.keys() {
                let e = &mut ext[at.sheet as usize];
                e.0 = e.0.max(at.row + 1);
                e.1 = e.1.max(at.col + 1);
            }
            for (s, (rows, cols)) in ext.into_iter().enumerate() {
                if rows > 0 && cols > 0 && Some(s as u16) != spec.skip_sheet {
                    query_rects.push(Rect::new(s as u16, 0, 0, rows - 1, cols - 1));
                }
            }
        }
    }
    // The names that can be read by name are also listed by `--meta`.
    let queryable_names: Vec<String> = named.iter().map(|b| b.name.clone()).collect();
    let queryable = match spec.query {
        QueryLevel::Declared => Queryable::Declared,
        QueryLevel::Named => Queryable::Named(named),
        QueryLevel::Any => Queryable::Any {
            excluded: spec.skip_sheet.into_iter().collect(),
            names: named,
        },
    };

    trace(&t0, "bindings");
    let n = formulas.len();
    let (ats, mut kinds): (Vec<CellRef>, Vec<CellKind>) = formulas.into_iter().unzip();
    let mut ids: FastMap<CellRef, u32> = FastMap::with_capacity_and_hasher(n, Default::default());
    let mut by_sheet: Vec<Vec<u32>> = vec![Vec::new(); model.sheets.len().max(1)];
    for (i, at) in ats.iter().enumerate() {
        ids.insert(*at, i as u32);
        by_sheet[at.sheet as usize].push(i as u32);
    }
    let index = Index {
        ats: &ats,
        ids: &ids,
        by_sheet: &by_sheet,
    };
    let view = ConstView::new(model, &cached);

    trace(&t0, "index built");
    // ----- 2. dependency graph, lazily, from the outputs -----
    let mut g = Lazy {
        index,
        kinds: &kinds,
        adj: vec![None; n],
        rect_ids: FastMap::default(),
        rect_members: Vec::new(),
        flags: FastMap::default(),
        deps: Vec::new(),
    };
    let mut roots: Vec<u32> = Vec::new();
    for r in outputs.iter().map(|b| &b.rect).chain(&query_rects) {
        g.members(r, &mut roots);
    }
    let mut visited = vec![false; n];
    let mut cone: Vec<u32> = Vec::new();
    let mut stack = roots;
    let mut all = false;
    while let Some(v) = stack.pop() {
        if g.seen(&mut visited, v) {
            continue;
        }
        cone.push(v);
        let adj = g.adj_of(v).to_vec();
        if (v as usize) < n && g.flags.get(&v).is_some_and(|f| f.dynamic) && !all {
            all = true;
            report.dynamic_refs = true;
            stack.extend(0..n as u32);
        }
        stack.extend(adj);
    }
    // Compact ids over the cone.
    let mut compact = vec![u32::MAX; visited.len()];
    for (i, &v) in cone.iter().enumerate() {
        compact[v as usize] = i as u32;
    }
    let adj_c: Vec<Vec<u32>> = cone
        .iter()
        .map(|&v| g.adj_of(v).iter().map(|&w| compact[w as usize]).collect())
        .collect();
    let csr = Csr::from_adjacency(&adj_c);
    drop(adj_c);
    trace(&t0, "cone + csr");
    let sccs = graph::tarjan(&csr);
    report.cells_reachable = cone.iter().filter(|&&v| (v as usize) < n).count();

    trace(&t0, "scc");
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
                        let f = g.flags.get(&v);
                        let _ = e;
                        t |= !spec.fold || f.is_some_and(|f| f.dynamic || f.volatile);
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

    trace(&t0, "taint");
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
            queryable: Queryable::Declared,
            meta: Vec::new(),
            diagnostics: Vec::new(),
            build: Default::default(),
        };
        let mut eng = Engine::with_base(&tmp, opts.clone(), &view);
        eng.run();
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
                        format!("{sheet}!{}", crate::addr::a1(a.row, a.col))
                    })
                    .collect();
                msgs.push(format!(
                    "{} formula(s) {what}, so the build cannot reproduce them: {}",
                    at.len(),
                    list.join(", ")
                ));
            }
            return Err(CompileError(msgs.join("; ")));
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

    trace(&t0, "fold evaluated");
    // What the surviving computation (and the outputs) still read: folded
    // formulas as constants, and the workbook's own constants.
    let mut need_rects: FastSet<Rect> = FastSet::default();
    for (i, &v) in cone.iter().enumerate() {
        if tainted[i]
            && (v as usize) < n
            && let CellKind::Formula(e) = &kinds[v as usize]
        {
            need_rects.extend(dependencies(e).0);
        }
    }
    for &id in &blank_ids {
        if let CellKind::Formula(e) = &kinds[id as usize] {
            need_rects.extend(dependencies(e).0);
        }
    }
    need_rects.extend(outputs.iter().map(|b| b.rect));
    need_rects.extend(query_rects.iter().copied());
    let mut need_folded: FastSet<u32> = FastSet::default();
    let mut const_cells: Vec<CellRef> = Vec::new();
    let mut members = Vec::new();
    for r in &need_rects {
        members.clear();
        index.members(r, &mut members);
        need_folded.extend(members.iter().copied().filter(|&id| {
            compact[id as usize] != u32::MAX && !tainted[compact[id as usize] as usize]
        }));
        view.consts_in(r, &ids, &mut const_cells);
    }
    if all {
        need_folded.extend(
            cone.iter()
                .copied()
                .filter(|&v| (v as usize) < n && !tainted[compact[v as usize] as usize]),
        );
        const_cells.clear();
        view.all_consts(&ids, &mut const_cells);
    }
    const_cells.sort_unstable();
    const_cells.dedup();

    trace(&t0, "needed");
    // ----- assemble -----
    let mut consts: Vec<(CellRef, Value)> = const_cells
        .iter()
        .map(|&at| (at, view.get(at).unwrap_or_default()))
        .collect();
    consts.extend(
        need_folded
            .iter()
            .filter(|id| !blank_ids.contains(id))
            .map(|&id| {
                let v = folded[id as usize].clone();
                (ats[id as usize], v.expect("a folded formula has its value"))
            }),
    );
    consts.sort_unstable_by_key(|(at, _)| *at);
    let mut cells: Vec<CellDef> = consts
        .into_iter()
        .map(|(at, v)| CellDef {
            at,
            kind: CellKind::Const(v),
        })
        .collect();
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
                Expr::Binary(sheetrs_engine::ir::BinOp::Concat, a, b) => {
                    Some(code_text(a, model, lookup)? + &code_text(b, model, lookup)?)
                }
                _ => None,
            }
        }
        let lookup = |r: CellRef| -> Option<Option<Value>> {
            let id = *ids.get(&r)?;
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
                        Some(t) if crate::format::has_foreign_language_tag(&t) => {
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
            queryable,
            meta: program_meta(model, spec, &queryable_names),
            diagnostics: diags,
            build: build_info(model, spec),
        },
        report,
    ))
}

/// The moment of a build in UTC (`2026-10-30T12:00:00Z`): `SOURCE_DATE_EPOCH` (seconds since 1970)
/// when it is set, so that a build can be repeated exactly; else the clock. A value that is not a
/// number of seconds is an error: a build that was meant to repeat must not silently use the clock.
pub fn build_time() -> Result<String, String> {
    let seconds = match std::env::var("SOURCE_DATE_EPOCH") {
        Ok(value) => value.trim().parse::<i64>().map_err(|_| {
            format!("SOURCE_DATE_EPOCH must be a number of seconds since 1970, not `{value}`")
        })?,
        Err(_) => std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map_or(0, |d| d.as_secs().cast_signed()),
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

/// How the program is built: the compiler, the options that shape it and the workbook. The time
/// is left empty: the command that builds sets it with [`build_time`].
fn build_info(model: &Model, spec: &Spec) -> BuildInfo {
    BuildInfo {
        compiler: format!("sheetrs-compiler {}", env!("CARGO_PKG_VERSION")),
        target: format!("{}-{}", std::env::consts::ARCH, std::env::consts::OS),
        built_at: String::new(),
        options: BuildOptions {
            outputs: match spec.query {
                QueryLevel::Declared => "declared",
                QueryLevel::Named => "named",
                QueryLevel::Any => "any",
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
/// of the settings sheet, and the defined names the program can read by name (`name`, only when it
/// was built with `outputs = named` or `any`).
fn program_meta(model: &Model, spec: &Spec, names: &[String]) -> Vec<Meta> {
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
    let mut names: Vec<&String> = names.iter().collect();
    names.sort();
    meta.extend(
        names
            .into_iter()
            .filter_map(|name| Some(entry(name, model.names.get(name)?, "name"))),
    );
    meta
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
    fn consts_in(&self, r: &Rect, nodes: &FastMap<CellRef, u32>, out: &mut Vec<CellRef>) {
        let total = self.model.cells.len() as u64;
        if r.area() <= total * 4 + 64 {
            for row in r.r0..=r.r1 {
                for col in r.c0..=r.c1 {
                    let at = CellRef::new(r.sheet, row, col);
                    if !nodes.contains_key(&at) && self.get(at).is_some() {
                        out.push(at);
                    }
                }
            }
        } else {
            out.extend(
                self.model
                    .cells
                    .keys()
                    .filter(|at| r.contains(*at) && !nodes.contains_key(at))
                    .filter(|at| self.get(*at).is_some()),
            );
        }
    }

    fn all_consts(&self, nodes: &FastMap<CellRef, u32>, out: &mut Vec<CellRef>) {
        out.extend(
            self.model
                .cells
                .keys()
                .filter(|at| !nodes.contains_key(at) && self.get(*at).is_some()),
        );
        out.extend(self.cached.keys().filter(|at| !nodes.contains_key(at)));
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

/// Where the cells that are graph nodes sit: node `i` is at `ats[i]`.
#[derive(Clone, Copy)]
struct Index<'a> {
    ats: &'a [CellRef],
    ids: &'a FastMap<CellRef, u32>,
    by_sheet: &'a [Vec<u32>],
}

impl Index<'_> {
    /// Existing cells inside `r`.
    fn members(&self, r: &Rect, out: &mut Vec<u32>) {
        let slots = &self.by_sheet[r.sheet as usize];
        if r.area() <= slots.len() as u64 * 4 + 64 {
            for row in r.r0..=r.r1 {
                for col in r.c0..=r.c1 {
                    if let Some(&id) = self.ids.get(&CellRef::new(r.sheet, row, col)) {
                        out.push(id);
                    }
                }
            }
        } else {
            out.extend(
                slots
                    .iter()
                    .copied()
                    .filter(|&s| r.contains(self.ats[s as usize])),
            );
        }
    }
}

/// Lazily-built dependency adjacency over cell nodes (`0..n`) and shared
/// rectangle nodes (`n..`).
struct Lazy<'a> {
    index: Index<'a>,
    kinds: &'a [CellKind],
    adj: Vec<Option<Vec<u32>>>,
    rect_ids: FastMap<Rect, u32>,
    rect_members: Vec<Vec<u32>>,
    flags: FastMap<u32, Flags>,
    /// The references of the formula being read, reused by every formula.
    deps: Vec<Rect>,
}

impl Lazy<'_> {
    fn n(&self) -> u32 {
        self.kinds.len() as u32
    }

    fn seen(&self, visited: &mut Vec<bool>, v: u32) -> bool {
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
                    self.flags.insert(v, flags);
                    let mut out = Vec::with_capacity(deps.len());
                    for &r in &deps {
                        // One cell is its own node: no rectangle to share.
                        if r.is_single() {
                            out.extend(self.index.ids.get(&r.top_left()).copied());
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

fn dependencies(e: &Expr) -> (Vec<Rect>, Flags) {
    let mut deps = Vec::new();
    let flags = dependencies_into(e, &mut deps);
    (deps, flags)
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
    use super::format_utc;

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
}
