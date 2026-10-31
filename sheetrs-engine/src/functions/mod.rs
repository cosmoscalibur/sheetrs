//! Function registry. Every function has a namespaced identity; the parser
//! asks for `name`
//! exists, the `core` implementation otherwise.

mod agg_ext;
mod arr_ext;
mod bahttext;
mod bessel;
mod complex;
mod convert;
mod date_ext;
mod datetime;
mod db_funcs;
mod dyn_array;
mod fin_ext;
mod finance;
mod forecast_ets;
mod fourier;
mod info;
mod logic;
mod lookup;
mod math;
mod math_ext;
mod matrix_lookup;
mod namespaces;
mod office;
mod pricing;
mod regex_fn;
mod regression;
mod securities;
mod special;
mod stat_dist;
mod stat_ext;
mod stats;
mod text;
mod text_bytes;
mod text_ext;
mod width_data;

use crate::engine::{Engine, Res};
use crate::ir::{Expr, FuncId, Ns};
use crate::value::Value;
use std::collections::HashMap;
use std::sync::LazyLock;

/// Scalar functions that act element-wise over arrays in array context.
const LIFTED: &[&str] = &[
    "LEN",
    "UPPER",
    "LOWER",
    "TRIM",
    "PROPER",
    "LEFT",
    "RIGHT",
    "MID",
    "ABS",
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "INT",
    "MOD",
    "SQRT",
    "EXP",
    "LN",
    "SIGN",
    "POWER",
    "ISNUMBER",
    "ISTEXT",
    "ISBLANK",
    "ISERROR",
    "ISNA",
    "ISEVEN",
    "ISODD",
    "VALUE",
    "TEXT",
    "SEARCH",
    "FIND",
    "SUBSTITUTE",
    "YEAR",
    "MONTH",
    "DAY",
    "WEEKDAY",
    "EXACT",
    "REPT",
    "CODE",
    "CHAR",
    "NOT",
    "CONCATENATE",
];

/// Functions whose arguments outside the domain are `#NUM!` in LibreOffice
/// (`Err:502` for the rest, which the readers translate to `#VALUE!`); found with
/// the oracle (`tests/assets/semantics.ods`).
pub const LO_DOMAIN_NUM: &[&str] = &["ASIN", "ACOS", "POWER"];

/// Arity limit meaning "any number of arguments".
pub const VAR: u8 = u8::MAX;

pub type EagerFn = fn(&mut Engine, &[Value]) -> Res<Value>;
/// Special forms receive their arguments unevaluated.
pub type LazyFn = fn(&mut Engine, &[Expr]) -> Res<Value>;

#[derive(Clone, Copy)]
pub enum Imp {
    Eager(EagerFn),
    Lazy(LazyFn),
}

pub struct FuncDef {
    pub ns: Ns,
    pub name: &'static str,
    pub min: u8,
    pub max: u8,
    pub imp: Imp,
    /// Result can change between evaluations (`RAND`, `NOW`): never folded.
    pub volatile: bool,
    /// May dereference cells unknown at build time (`INDIRECT`, `OFFSET`).
    pub dynamic: bool,
    /// Arguments are evaluated in array context: operators over ranges work
    /// element-wise instead of intersecting with the current cell.
    pub array_args: bool,
    /// Scalar function applied element-wise when given arrays in array context.
    pub lift: bool,
    /// Arguments that spread in array context (bit per position); 0 = all.
    pub lift_mask: u32,
}

pub(crate) struct Reg(Vec<FuncDef>);

impl Reg {
    pub(crate) fn eager(
        &mut self,
        name: &'static str,
        min: u8,
        max: u8,
        f: EagerFn,
    ) -> &mut FuncDef {
        self.push(name, min, max, Imp::Eager(f))
    }

    pub(crate) fn lazy(&mut self, name: &'static str, min: u8, max: u8, f: LazyFn) -> &mut FuncDef {
        self.push(name, min, max, Imp::Lazy(f))
    }

    /// Register `name` in a namespace other than the one the catalogue gives it.
    pub(crate) fn eager_in(
        &mut self,
        ns: Ns,
        name: &'static str,
        min: u8,
        max: u8,
        f: EagerFn,
    ) -> &mut FuncDef {
        let def = self.push(name, min, max, Imp::Eager(f));
        def.ns = ns;
        def
    }

    fn push(&mut self, name: &'static str, min: u8, max: u8, imp: Imp) -> &mut FuncDef {
        self.0.push(FuncDef {
            ns: Ns::Core,
            name,
            min,
            max,
            imp,
            volatile: false,
            dynamic: false,
            array_args: false,
            lift: false,
            lift_mask: 0,
        });
        let last = self.0.len() - 1;
        &mut self.0[last]
    }
}

impl FuncDef {
    pub(crate) fn volatile(&mut self) -> &mut Self {
        self.volatile = true;
        self
    }

    pub(crate) fn array_args(&mut self) -> &mut Self {
        self.array_args = true;
        self
    }

    pub(crate) fn dynamic(&mut self) -> &mut Self {
        self.dynamic = true;
        self
    }
}

pub fn table() -> &'static [FuncDef] {
    &REGISTRY.0
}

struct Registry(Vec<FuncDef>, HashMap<(Ns, &'static str), u16>);

static REGISTRY: LazyLock<Registry> = LazyLock::new(|| {
    let mut r = Reg(Vec::new());
    math::register(&mut r);
    math_ext::register(&mut r);
    stats::register(&mut r);
    stat_ext::register(&mut r);
    stat_dist::register(&mut r);
    logic::register(&mut r);
    text::register(&mut r);
    text_ext::register(&mut r);
    lookup::register(&mut r);
    arr_ext::register(&mut r);
    dyn_array::register(&mut r);
    db_funcs::register(&mut r);
    regression::register(&mut r);
    securities::register(&mut r);
    text_bytes::register(&mut r);
    bahttext::register(&mut r);
    bessel::register(&mut r);
    complex::register(&mut r);
    convert::register(&mut r);
    forecast_ets::register(&mut r);
    info::register(&mut r);
    fourier::register(&mut r);
    office::register(&mut r);
    pricing::register(&mut r);
    regex_fn::register(&mut r);
    agg_ext::register(&mut r);
    datetime::register(&mut r);
    date_ext::register(&mut r);
    finance::register(&mut r);
    fin_ext::register(&mut r);
    for d in &mut r.0 {
        // The namespace of a function registered without one comes from the catalogue.
        if d.ns == Ns::Core
            && let Some(&(_, ns)) = namespaces::OUTSIDE_CORE.iter().find(|(n, _)| *n == d.name)
        {
            d.ns = ns;
        }
        d.lift = LIFTED.contains(&d.name);
        // The criteria of the conditional aggregates spread; their ranges do not.
        d.lift_mask = match d.name {
            "COUNTIF" | "SUMIF" | "AVERAGEIF" => 1 << 1,
            "COUNTIFS" => 0b10_1010_1010,
            "SUMIFS" | "AVERAGEIFS" | "MAXIFS" | "MINIFS" => 0b101_0101_0100,
            _ => 0,
        };
        d.lift |= d.lift_mask != 0;
    }
    let map =
        r.0.iter()
            .enumerate()
            .map(|(i, d)| ((d.ns, d.name), i as u16))
            .collect();
    Registry(r.0, map)
});

/// The function `name` (uppercase) of exactly this namespace.
pub fn lookup_in(ns: Ns, name: &str) -> Option<FuncId> {
    REGISTRY.1.get(&(ns, name)).map(|&i| FuncId(i))
}

/// Resolve `name` (uppercase) written in `ns`: that namespace first, then `core`, then any other
/// (a formula that leaves out a prefix the function needs still finds it). `None`: unknown.
pub fn lookup(ns: Ns, name: &str) -> Option<FuncId> {
    lookup_in(ns, name)
        .or_else(|| lookup_in(Ns::Core, name))
        .or_else(|| {
            [Ns::Microsoft, Ns::OpenOffice, Ns::LibreOffice, Ns::Legacy]
                .into_iter()
                .find_map(|n| lookup_in(n, name))
        })
}

/// The namespace an XLSX gives a name it writes without `_xlfn.`, where that is not `core`
/// (`CEILING` is Excel's, `FDIST` the legacy one).
pub fn xlsx_bare(name: &str) -> Option<Ns> {
    namespaces::XLSX_BARE
        .iter()
        .find(|(n, _)| *n == name)
        .map(|&(_, ns)| ns)
}

/// Whether LibreOffice has a function of this name (uppercase, without prefix): see
/// `namespaces::CATALOGUE`. A function it has and the engine does not implement stops the build.
pub fn in_catalogue(name: &str) -> bool {
    namespaces::CATALOGUE.binary_search(&name).is_ok()
}

// ----- argument helpers shared by the function modules -----

/// Argument `i` unless it is absent or was omitted (`F(a,,c)`).
pub(crate) fn opt(args: &[Value], i: usize) -> Option<&Value> {
    args.get(i).filter(|v| !matches!(v, Value::Empty))
}

pub(crate) fn opt_num(e: &mut Engine, args: &[Value], i: usize, default: f64) -> Res<f64> {
    match opt(args, i) {
        Some(v) => e.num(v),
        None => Ok(default),
    }
}

pub(crate) fn num(n: f64) -> Res<Value> {
    if n.is_finite() {
        Ok(Value::Number(n))
    } else {
        Err(crate::value::ErrorKind::Num)
    }
}

/// Round half away from zero at `digits` decimal places on the 15 significant
/// digits of `x` (the way spreadsheets round).
pub fn round_decimal(x: f64, digits: i32) -> f64 {
    math::round_dec(x, digits, math::Mode::Nearest)
}
