//! Verification against a workbook's own cached values.
//!
//! A workbook saved by Excel or LibreOffice holds, next to every formula, the
//! value the application computed. Recomputing the formulas with the engine, with
//! the inputs the file holds (its defaults), and comparing is the check of
//! fidelity on a real model.

use crate::addr::{self, CellRef, Rect};
use crate::compile::{BindingSpec, CompileError, Spec, compile};
use crate::engine::Engine;
use crate::model::Model;
use crate::run::Overrides;
use crate::value::{ErrorKind, Value};
use std::collections::BTreeSet;

#[derive(Debug, Clone)]
pub struct Options {
    /// Sheets to check; all when empty.
    pub sheets: Vec<String>,
    /// Relative tolerance for numbers.
    pub tolerance: f64,
    /// Differences below this are noise whatever the magnitude: a balance that
    /// should be zero ends at 1e-9 or -1e-8 depending on the order of the sums.
    pub abs_tolerance: f64,
}

impl Default for Options {
    fn default() -> Self {
        Self {
            sheets: Vec::new(),
            tolerance: 1e-9,
            abs_tolerance: 1e-6,
        }
    }
}

#[derive(Debug, Clone)]
pub struct Mismatch {
    pub at: String,
    pub formula: String,
    pub expected: String,
    pub got: String,
    pub kind: &'static str,
}

#[derive(Debug, Clone, Default)]
pub struct SheetReport {
    pub name: String,
    pub formulas: usize,
    pub matched: usize,
    /// Formulas with `NOW`, `TODAY` or `RAND`: their cached value is from another day.
    pub volatile: usize,
    /// Formulas whose syntax the compiler keeps as the cached value.
    pub unparsed: usize,
    pub mismatches: Vec<Mismatch>,
}

#[derive(Debug, Default)]
pub struct Verification {
    pub sheets: Vec<SheetReport>,
    /// Functions the engine does not implement.
    pub unsupported: BTreeSet<String>,
    pub diagnostics: Vec<String>,
}

impl Verification {
    pub fn mismatches(&self) -> usize {
        self.sheets.iter().map(|s| s.mismatches.len()).sum()
    }
}

fn show(v: &Value) -> String {
    match v {
        Value::Empty => "(blank)".into(),
        Value::Number(n) => format!("{n}"),
        Value::Text(t) => format!("{t:?}"),
        Value::Bool(b) => b.to_string().to_uppercase(),
        Value::Error(e) => e.as_str().to_string(),
        _ => "(array)".into(),
    }
}

/// `None` when `got` is what `expected` says; else what differs.
fn differs(expected: &Value, got: &Value, tol: f64, abs: f64) -> Option<&'static str> {
    let blank =
        |v: &Value| matches!(v, Value::Empty) || matches!(v, Value::Text(t) if t.is_empty());
    match (expected, got) {
        (Value::Number(a), Value::Number(b)) => {
            let limit = abs.max(tol * a.abs().max(b.abs()));
            ((a - b).abs() > limit).then_some("number")
        }
        (Value::Text(a), Value::Text(b)) => (a != b).then_some("text"),
        (Value::Bool(a), Value::Bool(b)) => (a != b).then_some("boolean"),
        // LibreOffice stores the result of TRUE()/FALSE() as the number 1/0 (with a
        // boolean format) where this engine holds a boolean: the same value.
        (Value::Number(n), Value::Bool(b)) | (Value::Bool(b), Value::Number(n))
            if *n == f64::from(u8::from(*b)) =>
        {
            None
        }
        // A circular reference is `Err:522` in LibreOffice, which files store as `#VALUE!`.
        (Value::Error(ErrorKind::Value), Value::Error(ErrorKind::Circ)) => None,
        (Value::Error(a), Value::Error(b)) => (a != b).then_some("error kind"),
        (a, b) if blank(a) && blank(b) => None,
        (Value::Empty, Value::Number(n)) | (Value::Number(n), Value::Empty) if *n == 0.0 => {
            Some("blank vs 0")
        }
        _ => Some("type"),
    }
}

fn is_volatile(formula: &str) -> bool {
    let f = formula.to_ascii_uppercase();
    ["NOW(", "TODAY(", "RAND(", "RANDBETWEEN("]
        .iter()
        .any(|k| f.contains(k))
}

/// Recompute every formula of the selected sheets and compare with the cached
/// values of the file. `spec` carries the options of the build (locale, folding).
///
/// `reference`: a copy of the workbook whose cached values are the ones to compare with
/// (what LibreOffice recalculated, see `recalc`); `None` compares with the values `model`
/// itself caches.
pub fn verify(
    model: &Model,
    spec: &Spec,
    opts: &Options,
    reference: Option<&Model>,
) -> Result<Verification, CompileError> {
    let selected: Vec<u16> = (0..model.sheets.len() as u16)
        .filter(|&i| {
            opts.sheets.is_empty()
                || opts
                    .sheets
                    .iter()
                    .any(|n| addr::same_name(n, &model.sheets[i as usize]))
        })
        .collect();
    if selected.is_empty() {
        return Err(CompileError(format!(
            "none of the sheets {:?} exists (sheets: {})",
            opts.sheets,
            model.sheets.join(", ")
        )));
    }

    // Formula cells per sheet, in row bands small enough to bind as outputs.
    let mut cells: Vec<(CellRef, &str)> = model
        .cells
        .iter()
        .filter(|(at, c)| c.formula.is_some() && selected.contains(&at.sheet))
        .map(|(at, c)| (at, c.formula.as_deref().unwrap_or("")))
        .collect();
    cells.sort_unstable_by_key(|(at, _)| *at);
    let mut outputs = Vec::new();
    let mut i = 0;
    while i < cells.len() {
        let sheet = cells[i].0.sheet;
        let mut j = i;
        let (mut r0, mut r1, mut c0, mut c1) = (u32::MAX, 0, u32::MAX, 0);
        while j < cells.len() && cells[j].0.sheet == sheet {
            let at = cells[j].0;
            let (nr0, nr1, nc0, nc1) = (
                r0.min(at.row),
                r1.max(at.row),
                c0.min(at.col),
                c1.max(at.col),
            );
            if u64::from(nr1 - nr0 + 1) * u64::from(nc1 - nc0 + 1) > 1_000_000 && j > i {
                break;
            }
            (r0, r1, c0, c1) = (nr0, nr1, nc0, nc1);
            j += 1;
        }
        let rect = Rect::new(sheet, r0, c0, r1, c1);
        outputs.push(BindingSpec::named(
            &format!("band{}", outputs.len()),
            &addr::qualified(&model.sheets, &rect),
        ));
        i = j;
    }

    let spec = Spec {
        inputs: Vec::new(),
        outputs,
        meta: Vec::new(),
        skip_sheet: None,
        inspect: true,
        ..spec.clone()
    };
    let (program, report) = compile(model, &spec)?;
    let mut engine = Engine::new(&program, program.effective_options(&Overrides::default()));
    engine.run();

    let unparsed: std::collections::HashSet<String> = program
        .diagnostics
        .iter()
        .filter(|d| d.contains("unsupported syntax"))
        .filter_map(|d| d.split(": ").next().map(str::to_string))
        .collect();

    let mut out = Verification {
        unsupported: report.unsupported,
        diagnostics: program.diagnostics.iter().take(20).cloned().collect(),
        ..Verification::default()
    };
    for &s in &selected {
        out.sheets.push(SheetReport {
            name: model.sheets[s as usize].clone(),
            ..SheetReport::default()
        });
    }
    for (at, formula) in cells {
        let pos = selected.iter().position(|&s| s == at.sheet).unwrap_or(0);
        let rep = &mut out.sheets[pos];
        rep.formulas += 1;
        let label = format!("{}!{}", rep.name, addr::a1(at.row, at.col));
        if is_volatile(formula) {
            rep.volatile += 1;
            continue;
        }
        if unparsed.contains(&label) {
            rep.unparsed += 1;
            continue;
        }
        let expected = match reference {
            None => model.cells.get(&at),
            Some(r) => r
                .sheet_id(&model.sheets[at.sheet as usize])
                .and_then(|s| r.cells.get(&CellRef::new(s, at.row, at.col))),
        }
        .map_or(Value::Empty, |c| c.value.clone());
        let got = engine.get(at);
        match differs(&expected, &got, opts.tolerance, opts.abs_tolerance) {
            None => rep.matched += 1,
            Some(kind) => rep.mismatches.push(Mismatch {
                at: label,
                formula: formula.chars().take(120).collect(),
                expected: show(&expected),
                got: show(&got),
                kind,
            }),
        }
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::ModelCell;

    fn cached(m: &mut Model, a1: &str, formula: &str, value: Value) {
        let at = addr::parse_area(a1)
            .map(|a| CellRef::new(0, a.r0, a.c0))
            .unwrap();
        m.cells.insert(
            at,
            ModelCell {
                value,
                formula: Some(formula.into()),
                array: false,
            },
        );
    }

    #[test]
    fn compares_with_the_cached_values() {
        let mut m = Model::new();
        let s = m.add_sheet("S");
        m.set(s, "A1", Value::Number(2.0));
        cached(&mut m, "B1", "A1*2", Value::Number(4.0)); // right
        cached(&mut m, "B2", "A1*3", Value::Number(7.0)); // wrong number
        cached(
            &mut m,
            "B3",
            "A1/0",
            Value::Error(crate::value::ErrorKind::Value),
        ); // wrong error
        cached(&mut m, "B4", "A9", Value::Number(0.0)); // blank reference holds 0
        cached(&mut m, "B5", "NOW()", Value::Number(1.0)); // not comparable
        cached(&mut m, "B6", "A1+0.1", Value::Number(2.1 + 1e-12)); // within tolerance
        let spec = Spec::default();
        let v = verify(&m, &spec, &Options::default(), None).unwrap();
        let r = &v.sheets[0];
        assert_eq!(
            (r.formulas, r.matched, r.volatile),
            (6, 3, 1),
            "{:?}",
            r.mismatches
        );
        let kinds: Vec<_> = r
            .mismatches
            .iter()
            .map(|m| (m.at.as_str(), m.kind))
            .collect();
        assert_eq!(kinds, [("S!B2", "number"), ("S!B3", "error kind")]);
        assert!(
            verify(
                &m,
                &spec,
                &Options {
                    sheets: vec!["nope".into()],
                    ..Options::default()
                },
                None
            )
            .is_err()
        );
    }
}
