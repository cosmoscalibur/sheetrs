//! Function registry. Every function has a namespaced identity; the parser
//! asks for `name`
//! exists, the `core` implementation otherwise.

mod agg_ext;
mod arity;
mod arr_ext;
mod bahttext;
mod bessel;
mod complex;
mod computed_error_table;
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
mod keep_format_table;
mod lift_table;
mod literal_error;
mod literal_error_table;
mod logic;
mod lookup;
mod marker_table;
mod math;
mod math_ext;
mod matrix_lookup;
mod namespaces;
mod office;
mod omitted_table;
mod pricing;
mod propagating_error_table;
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
mod threed;
mod threed_table;
mod width_data;

pub use literal_error::{
    computed_error_first, ets_series, is_criterion, literal_answer, propagated_error,
    token_as_array, wants_reference,
};
pub use threed::{Outcome, is_variadic, outcome};

/// What a function makes of an argument that is written empty (`F(a,,c)`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Omitted {
    /// As if it was not there: the engine's own default for the argument.
    Absent,
    /// Not written at all, when it is the last (`DROP(r,)` is `DROP(r)`); the default where it is not.
    Dropped,
    /// The number 0.
    Zero,
    /// An empty text.
    EmptyText,
    /// The function gives this error.
    Error(ErrorKind),
}

/// What the function `name` of the namespace `ns` makes of its argument number `arg` (from 0) being
/// written empty, as `LibreOffice` does (`scripts/probe_omitted.py` asks it of every function).
#[must_use]
pub fn omitted(name: &str, ns: Ns, arg: usize) -> Omitted {
    let Ok(i) =
        omitted_table::OMITTED.binary_search_by(|&(n, s, _)| (n, s).cmp(&(name, ns.as_str())))
    else {
        return Omitted::Absent;
    };
    let Some(&(_, what)) = omitted_table::OMITTED[i]
        .2
        .iter()
        .find(|&&(p, _)| usize::from(p) == arg)
    else {
        return Omitted::Absent;
    };
    match what {
        "Z" => Omitted::Zero,
        "D" => Omitted::Dropped,
        "E" => Omitted::EmptyText,
        "!v" => Omitted::Error(ErrorKind::Value),
        "!a" => Omitted::Error(ErrorKind::Na),
        "!n" => Omitted::Error(ErrorKind::Name),
        "!r" => Omitted::Error(ErrorKind::Ref),
        "!d" => Omitted::Error(ErrorKind::DivZero),
        "!u" => Omitted::Error(ErrorKind::Num),
        code => code
            .strip_prefix('!')
            .and_then(|c| c.parse().ok())
            .map_or(Omitted::Absent, |c| Omitted::Error(ErrorKind::Libre(c))),
    }
}

use crate::engine::{Engine, Res};
use crate::ir::{Expr, FuncId, Ns};
use crate::value::{ErrorKind, Value};
use std::collections::HashMap;
use std::sync::LazyLock;

/// Functions whose arguments outside the domain are `#NUM!` in `LibreOffice`
/// (`Err:502` for the rest, which the readers translate to `#VALUE!`); found with
/// the oracle (`tests/assets/semantics.ods`).
pub const LO_DOMAIN_NUM: &[&str] = &["ASIN", "ACOS", "POWER"];

/// What `LibreOffice` makes of the error a function gives itself (an argument outside its domain, a
/// value of the wrong kind), for the functions the oracle (`tests/assets/semantics.ods`) says enough
/// of: its code (`Err:502` invalid argument, `Err:504` parameter list, `Err:523` the iteration does
/// not converge, …), or the plain Excel error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LoError {
    /// Every error of the function is this one: an error with its code (a file format without
    /// codes holds it as `#VALUE!`).
    Libre(u16),
    /// Every error of the function is the plain `#VALUE!` (`Err:519`).
    Value,
    /// Every error of the function is the plain `#NUM!` (`Err:503`).
    Num,
    /// The function has several errors, and what it raises outside the domain of its arguments is
    /// this code (the commonest one the function gives): the engine does not tell the other
    /// situations apart, so only the errors outside the domain are given the code.
    DomainIs(u16),
}

/// Per function, the error `LibreOffice` gives for what the function raises itself, from the
/// oracle's error cells (7,354 of them): 97 functions give `Err:502` throughout, 7 `Err:504`, and
/// the rest of what the oracle has is listed with the single code of its arguments outside the
/// domain. Sorted by name. A function that is not here keeps the generic rule: outside the domain is
/// `#NUM!` for [`LO_DOMAIN_NUM`] and `Err:502` for the rest, which is what most of them are, and a
/// plain `#VALUE!` stays plain. What is not known of a function with several codes (`INDEX` has
/// `Err:502` and `Err:504`) is a gap that `oracle::ods_matches_libreoffice` leaves unchecked: it
/// only stops a code that `LibreOffice` never gives for the function.
pub const LO_ERRORS: &[(&str, LoError)] = &[
    ("ACCRINTM", LoError::Libre(502)),
    ("ACOS", LoError::Num),
    ("ACOSH", LoError::Libre(502)),
    ("ACOTH", LoError::Libre(502)),
    ("AGGREGATE", LoError::Libre(502)),
    ("ARABIC", LoError::Libre(502)),
    ("AREAS", LoError::DomainIs(504)),
    ("ASIN", LoError::Num),
    ("ATANH", LoError::Libre(502)),
    ("B", LoError::Libre(502)),
    ("BAHTTEXT", LoError::DomainIs(511)),
    ("BASE", LoError::Libre(502)),
    ("BESSELK", LoError::Libre(502)),
    ("BESSELY", LoError::Libre(502)),
    ("BETA.DIST", LoError::Libre(502)),
    ("BETADIST", LoError::Libre(502)),
    ("BETAINV", LoError::Libre(502)),
    ("BINOMDIST", LoError::Libre(502)),
    ("BITAND", LoError::Libre(502)),
    ("CHAR", LoError::Libre(502)),
    ("CHOOSECOLS", LoError::Libre(504)),
    ("CHOOSEROWS", LoError::Libre(504)),
    ("COMBIN", LoError::Libre(502)),
    ("CONFIDENCE", LoError::Libre(502)),
    ("COSH", LoError::Num),
    ("COUPDAYBS", LoError::DomainIs(502)),
    ("COUPDAYS", LoError::DomainIs(502)),
    ("COUPDAYSNC", LoError::DomainIs(502)),
    ("COUPNCD", LoError::DomainIs(502)),
    ("COUPPCD", LoError::DomainIs(502)),
    ("CRITBINOM", LoError::Libre(502)),
    ("CUMIPMT", LoError::Libre(502)),
    ("CURRENT", LoError::Libre(518)),
    ("DATE", LoError::Libre(502)),
    ("DATEVALUE", LoError::Libre(502)),
    ("DAVERAGE", LoError::DomainIs(504)),
    ("DAYSINMONTH", LoError::Value),
    ("DCOUNT", LoError::DomainIs(504)),
    ("DCOUNTA", LoError::DomainIs(504)),
    ("DEC2BIN", LoError::Libre(502)),
    ("DECIMAL", LoError::Libre(502)),
    ("DGET", LoError::DomainIs(504)),
    ("DISC", LoError::Libre(502)),
    ("DMAX", LoError::DomainIs(504)),
    ("DMIN", LoError::DomainIs(504)),
    ("DPRODUCT", LoError::DomainIs(504)),
    ("DROP", LoError::Libre(502)),
    ("DSTDEV", LoError::DomainIs(504)),
    ("DSTDEVP", LoError::DomainIs(504)),
    ("DSUM", LoError::DomainIs(504)),
    ("DURATION", LoError::Libre(502)),
    ("DVAR", LoError::DomainIs(504)),
    ("DVARP", LoError::DomainIs(504)),
    ("ENCODEURL", LoError::Value),
    ("EXP", LoError::Num),
    ("EXPAND", LoError::Libre(502)),
    ("F.DIST", LoError::Libre(502)),
    ("F.INV", LoError::Libre(502)),
    ("F.INV.RT", LoError::Libre(502)),
    ("FACTDOUBLE", LoError::Libre(502)),
    ("FIND", LoError::Value),
    ("FINV", LoError::Libre(502)),
    ("FISHER", LoError::Libre(502)),
    ("FORECAST.ETS.SEASONALITY", LoError::Value),
    ("FOURIER", LoError::DomainIs(502)),
    ("FTEST", LoError::Value),
    ("GAMMA.INV", LoError::Libre(502)),
    ("GAMMALN", LoError::Libre(502)),
    ("GAMMALN.PRECISE", LoError::Libre(502)),
    ("GCD", LoError::Libre(502)),
    ("GEOMEAN", LoError::Libre(502)),
    ("HARMEAN", LoError::Libre(502)),
    ("HEX2DEC", LoError::Libre(502)),
    ("HYPGEOMDIST", LoError::Libre(502)),
    ("INTRATE", LoError::Libre(502)),
    ("IPMT", LoError::Libre(502)),
    ("ISEVEN", LoError::Libre(504)),
    ("ISODD", LoError::Libre(504)),
    ("LARGE", LoError::Value),
    ("LCM", LoError::Libre(502)),
    ("LN", LoError::Libre(502)),
    ("LOG", LoError::Libre(502)),
    ("LOGINV", LoError::Libre(502)),
    ("LOGNORM.DIST", LoError::Libre(502)),
    ("MDETERM", LoError::Value),
    ("MIDB", LoError::Libre(502)),
    ("MODE", LoError::Value),
    ("MULTINOMIAL", LoError::Libre(502)),
    ("NEGBINOM.DIST", LoError::Libre(502)),
    ("NEGBINOMDIST", LoError::Libre(502)),
    ("NORMDIST", LoError::Libre(502)),
    ("NORMINV", LoError::Libre(502)),
    ("NORMSINV", LoError::Value),
    ("NUMBERVALUE", LoError::Value),
    ("OPT_PROB_HIT", LoError::Libre(504)),
    ("OPT_PROB_INMONEY", LoError::Libre(502)),
    ("OPT_TOUCH", LoError::Libre(502)),
    ("PERCENTILE.EXC", LoError::Libre(504)),
    ("PERCENTILE.INC", LoError::Libre(502)),
    ("PERCENTRANK", LoError::Value),
    ("PERMUT", LoError::Libre(502)),
    ("POISSON", LoError::Libre(502)),
    ("POWER", LoError::Num),
    ("PRICEDISC", LoError::Libre(502)),
    ("PRICEMAT", LoError::Libre(502)),
    ("QUARTILE", LoError::Libre(502)),
    ("QUARTILE.EXC", LoError::Libre(502)),
    ("QUOTIENT", LoError::Libre(502)),
    ("RANDARRAY", LoError::Libre(502)),
    ("RANDBETWEEN", LoError::Libre(502)),
    ("RECEIVED", LoError::Libre(502)),
    ("REPLACEB", LoError::Libre(502)),
    ("RIGHTB", LoError::Libre(502)),
    ("ROMAN", LoError::Libre(502)),
    ("SEARCH", LoError::DomainIs(502)),
    ("SEARCHB", LoError::DomainIs(502)),
    ("SHEETS", LoError::DomainIs(504)),
    ("SINH", LoError::Num),
    ("SMALL", LoError::Value),
    ("SORT", LoError::Libre(504)),
    ("SORTBY", LoError::Libre(504)),
    ("SQRT", LoError::Libre(502)),
    ("SQRTPI", LoError::Libre(502)),
    ("SUBTOTAL", LoError::Libre(502)),
    ("SUMPRODUCT", LoError::Libre(502)),
    ("T.DIST", LoError::Libre(502)),
    ("T.DIST.2T", LoError::Libre(502)),
    ("T.INV", LoError::Libre(502)),
    ("TBILLEQ", LoError::Libre(502)),
    ("TBILLPRICE", LoError::Libre(502)),
    ("TBILLYIELD", LoError::Libre(502)),
    ("TDIST", LoError::Libre(502)),
    ("TEXT", LoError::Libre(502)),
    ("TEXTAFTER", LoError::DomainIs(504)),
    ("TEXTBEFORE", LoError::DomainIs(504)),
    ("TEXTJOIN", LoError::DomainIs(504)),
    ("TIMEVALUE", LoError::Libre(502)),
    ("TINV", LoError::Libre(502)),
    ("TOCOL", LoError::DomainIs(504)),
    ("TRIMMEAN", LoError::Libre(502)),
    ("TYPE", LoError::Value),
    ("UNICODE", LoError::Libre(504)),
    ("VALUE", LoError::Libre(502)),
    ("WEEKNUM", LoError::Libre(502)),
    ("WORKDAY.INTL", LoError::Value),
    ("XMATCH", LoError::DomainIs(504)),
    ("YEARFRAC", LoError::Libre(502)),
    ("YIELD", LoError::DomainIs(502)),
    ("YIELDDISC", LoError::Libre(502)),
    ("YIELDMAT", LoError::Libre(502)),
];

/// The error `LibreOffice` gives for what `function` raises, when its errors are all the same.
#[must_use]
pub fn lo_error(function: &str) -> Option<LoError> {
    LO_ERRORS
        .binary_search_by(|(name, _)| (*name).cmp(function))
        .ok()
        .map(|i| LO_ERRORS[i].1)
}

/// What `LibreOffice` does with a call of `given` arguments of `def`, from the table of
/// `arity.rs` (probed for every function): `Some(code)` it refuses the count with that error
/// (`Err:504`, `Err:508`, `Err:509`, `Err:510`, `Err:511`, `Err:518`), `None` it takes it (or the function is not in the table).
#[must_use]
pub fn lo_arity_refusal(def: &FuncDef, given: usize) -> Option<ErrorKind> {
    let counts = def.lo_arity.as_bytes();
    let code = *counts.get(given).or_else(|| counts.last())?;
    match code {
        b'4' => Some(ErrorKind::Libre(504)),
        b'8' => Some(ErrorKind::Libre(508)),
        b'9' => Some(ErrorKind::Libre(509)),
        b'0' => Some(ErrorKind::Libre(510)),
        b'1' => Some(ErrorKind::Libre(511)),
        b'b' => Some(ErrorKind::Libre(518)),
        _ => None,
    }
}

/// The error of a call of `def` with `given` arguments that the engine cannot run: the code
/// `LibreOffice` gives when it refuses the count too; when it takes it (a call the engine does not
/// implement, like `SUM()`), the error value of a case that is not reproduced (the build stops
/// when it meets it). A function that is not in the table has no probe: too many arguments are
/// `Err:504` and too few `Err:511`.
#[must_use]
pub fn arity_error(def: &FuncDef, given: usize) -> ErrorKind {
    if let Some(refused) = lo_arity_refusal(def, given) {
        return refused;
    }
    if !def.lo_arity.is_empty() {
        crate::notes::note(crate::notes::ARITY);
        return ErrorKind::Unsupported;
    }
    ErrorKind::Libre(if given < def.min as usize { 511 } else { 504 })
}

/// Whether a call of `given` arguments of `def` has to be looked at by [`arity_error`]: the engine
/// does not take the count, or `LibreOffice` refuses a count the engine takes.
#[must_use]
pub fn arity_refused(def: &FuncDef, given: usize) -> bool {
    given < def.min as usize
        || given > def.max as usize
        || (def.lo_irregular && lo_arity_refusal(def, given).is_some())
}

/// The functions that take the arrays they are given whole, though the call over an array of two
/// numbers has two rows (what `scripts/probe_lift.py` takes for a function called for each element):
/// `VSTACK`, `MMULT`, `MODE.MULT`, the tests of two series…
const WHOLE_ARRAY: &[&str] = &[
    "CHITEST",
    "CHISQ.TEST",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "EXPAND",
    "F.TEST",
    "FTEST",
    "HSTACK",
    "MDETERM",
    "MINVERSE",
    "MMULT",
    "MODE.MULT",
    "T.TEST",
    "TTEST",
    "TOCOL",
    "TOROW",
    "TRANSPOSE",
    "VSTACK",
    "WRAPCOLS",
    "WRAPROWS",
];

/// The functions whose arguments `LibreOffice` evaluates as arrays in a formula that is not an array
/// formula, though their registration does not say so: `CORREL(A1:A9*1, B1:B9)` is the correlation
/// of the arrays, not of the two cells the formula meets. A bit per argument position (probed by
/// `scripts/probe_intersection.py`). The newer names (`F.TEST`, `T.TEST`, `CHITEST`,
/// `CHISQ.TEST`, `FORECAST.LINEAR`) are not on the list: there it meets the cell, as in any other function.
const FORCE_ARRAY: &[(&str, u32)] = &[
    ("CORREL", 0b11),
    ("PEARSON", 0b11),
    ("RSQ", 0b11),
    ("SLOPE", 0b11),
    ("INTERCEPT", 0b11),
    ("STEYX", 0b11),
    ("COVAR", 0b11),
    ("COVARIANCE.P", 0b11),
    ("COVARIANCE.S", 0b11),
    ("FTEST", 0b11),
    ("TTEST", 0b11),
    ("FORECAST", 0b110),
    ("PROB", 0b11),
    ("MODE", u32::MAX),
    ("MODE.SNGL", u32::MAX),
    // The condition of `FILTER` and the arrays of `XLOOKUP` and `XMATCH`, not what they look for.
    ("FILTER", 0b10),
    ("XLOOKUP", 0b110),
    ("XMATCH", 0b10),
    ("MATCH", 0b10),
    ("VLOOKUP", 0b10),
    ("HLOOKUP", 0b10),
    ("MDETERM", u32::MAX),
    ("MINVERSE", u32::MAX),
    ("MMULT", u32::MAX),
];

/// The functions that take an array but whose argument `LibreOffice` evaluates as in any other
/// function when the formula is not an array formula: an operator over a range is met by the cell of
/// the formula there, so `SORT(A1:A9*2)` is the sort of one number (`Err:504`, a value where an
/// array goes) and `INDEX(A1:A9*2, 1)` is `#REF!`. In an array formula they see the whole range.
const NOT_FORCED: &[&str] = &[
    "INDEX",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "TAKE",
    "DROP",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "HSTACK",
    "VSTACK",
    "TOCOL",
    "TOROW",
    "WRAPROWS",
    "WRAPCOLS",
    "EXPAND",
    "TREND",
    "GROWTH",
    "SERIESSUM",
];

/// Arguments that must be a range or an array: `LibreOffice` refuses a plain value there with the
/// error of the function (`INDEX(5, 1)` is `#REF!`, `MATCH(1, 5, 0)` is `Err:504`,
/// `SUMIFS(5, A1:A3, ">1")` is `Err:502`). Per function, the arguments (a bit each, the first is bit 0)
/// and the error. Found by asking `LibreOffice` for each function with a plain number in each position
/// that takes a range (`scripts/probe_scalar_args.py` lists the candidates). A cell reference, an array
/// written in the formula and the result of a function that gives an array are not plain values.
pub const MATRIX_ARGS: &[(&str, &[(u32, ErrorKind)])] = &[
    ("AVERAGEIF", &[(0b101, ErrorKind::Libre(504))]),
    (
        "AVERAGEIFS",
        &[
            (0b1010, ErrorKind::Libre(504)),
            (0b1, ErrorKind::Libre(502)),
        ],
    ),
    ("CHOOSECOLS", &[(0b1, ErrorKind::Libre(504))]),
    ("CHOOSEROWS", &[(0b1, ErrorKind::Libre(504))]),
    ("COLUMN", &[(0b1, ErrorKind::Libre(504))]),
    ("COUNTBLANK", &[(0b1, ErrorKind::Libre(504))]),
    ("COUNTIF", &[(0b1, ErrorKind::Libre(504))]),
    ("COUNTIFS", &[(0b01_0101_0101_0101, ErrorKind::Libre(504))]),
    ("DROP", &[(0b1, ErrorKind::Libre(504))]),
    ("EXPAND", &[(0b1, ErrorKind::Libre(504))]),
    ("FILTER", &[(0b10, ErrorKind::Libre(504))]),
    ("HLOOKUP", &[(0b10, ErrorKind::Libre(504))]),
    ("INDEX", &[(0b1, ErrorKind::Ref)]),
    ("MATCH", &[(0b10, ErrorKind::Libre(504))]),
    (
        "MAXIFS",
        &[
            (0b1010, ErrorKind::Libre(504)),
            (0b1, ErrorKind::Libre(502)),
        ],
    ),
    (
        "MINIFS",
        &[
            (0b1010, ErrorKind::Libre(504)),
            (0b1, ErrorKind::Libre(502)),
        ],
    ),
    ("OFFSET", &[(0b1, ErrorKind::Libre(504))]),
    ("ROW", &[(0b1, ErrorKind::Libre(504))]),
    ("SORT", &[(0b1, ErrorKind::Libre(504))]),
    ("SORTBY", &[(0b10, ErrorKind::Libre(504))]),
    ("SUMIF", &[(0b101, ErrorKind::Libre(504))]),
    (
        "SUMIFS",
        &[
            (0b1010, ErrorKind::Libre(504)),
            (0b1, ErrorKind::Libre(502)),
        ],
    ),
    ("TAKE", &[(0b1, ErrorKind::Libre(504))]),
    ("TOROW", &[(0b1, ErrorKind::Libre(504))]),
    ("UNIQUE", &[(0b1, ErrorKind::Libre(504))]),
    ("VLOOKUP", &[(0b10, ErrorKind::Libre(504))]),
    ("WRAPCOLS", &[(0b1, ErrorKind::Libre(504))]),
    ("WRAPROWS", &[(0b1, ErrorKind::Libre(504))]),
    ("XLOOKUP", &[(0b110, ErrorKind::Libre(504))]),
    ("XMATCH", &[(0b10, ErrorKind::Libre(504))]),
];

/// The error `LibreOffice` gives when one of the arguments of `def` that must be a range or an
/// array is a plain value (see [`MATRIX_ARGS`]).
#[must_use]
pub fn scalar_refused(def: &FuncDef, args: &[Value]) -> Option<ErrorKind> {
    // `TAKE(x)` and `DROP(x)` have no rows or columns to give: `Err:502` whatever `x` is.
    if args.len() == 1 && matches!(def.name, "TAKE" | "DROP") {
        return None;
    }
    def.matrix.iter().find_map(|&(mask, kind)| {
        args.iter()
            .enumerate()
            .any(|(i, v)| {
                i < 32
                    && mask >> i & 1 == 1
                    && matches!(v, Value::Number(_) | Value::Text(_) | Value::Bool(_))
            })
            .then_some(kind)
    })
}

/// The arguments (a bit per position) where a range of several cells is `Err:502` in `LibreOffice`
/// outside an array formula, instead of being met by the cell of the formula as in any other
/// argument: the second argument of the base conversions, of `DELTA`, `GESTEP` and `ERF`, and the
/// basis of `YEARFRAC` (the functions of its Analysis add-in take only the first argument that way).
const RANGE_IS_ERROR: &[(&str, u32)] = &[
    ("BIN2HEX", 0b10),
    ("BIN2OCT", 0b10),
    ("DEC2BIN", 0b10),
    ("DEC2HEX", 0b10),
    ("DEC2OCT", 0b10),
    ("HEX2BIN", 0b10),
    ("HEX2OCT", 0b10),
    ("OCT2BIN", 0b10),
    ("OCT2HEX", 0b10),
    ("DELTA", 0b10),
    ("GESTEP", 0b10),
    ("ERF", 0b10),
    ("YEARFRAC", 0b100),
];

/// The error `LibreOffice` gives when one of the arguments of `def` in `RANGE_IS_ERROR` is a range
/// of several cells.
#[must_use]
pub fn range_refused(def: &FuncDef, args: &[Value]) -> Option<ErrorKind> {
    (def.range_error != 0
        && args.iter().enumerate().any(|(i, v)| {
            i < 32 && def.range_error >> i & 1 == 1 && matches!(v, Value::Ref(r) if !r.is_single())
        }))
    .then_some(ErrorKind::Libre(502))
}

/// Arity limit meaning "any number of arguments".
pub const VAR: u8 = u8::MAX;

pub type EagerFn = fn(&mut Engine, &[Value]) -> Res<Value>;
/// Special forms receive their arguments unevaluated.
pub type LazyFn = fn(&mut Engine, &[Expr]) -> Res<Value>;

#[derive(Debug, Clone, Copy)]
pub enum Imp {
    Eager(EagerFn),
    Lazy(LazyFn),
}

// One flag per behaviour of a function that the compiler and the engine read separately.
#[allow(clippy::struct_excessive_bools)]
#[derive(Debug)]
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
    /// The arguments (a bit per position; all of them for [`u32::MAX`]) evaluated in array context
    /// even in a formula that is not an array formula: operators over ranges work element-wise
    /// there instead of meeting the cell of the formula. It is what `LibreOffice` does for the
    /// arguments it takes as arrays (see `FORCE_ARRAY` and `NOT_FORCED`).
    pub force_array: u32,
    /// The arguments where a range of several cells is `Err:502` (see `RANGE_IS_ERROR`).
    pub range_error: u32,
    /// Whether an error a calculation gave in any argument is the answer, before the arguments are checked
    /// (`ACCRINT("a","a","a","a","a","a",NA())` is `#N/A`): `propagating_error_table.rs`.
    pub propagates: bool,
    /// Scalar function applied element-wise when given arrays in array context.
    pub lift: bool,
    /// Arguments that spread in array context (bit per position); 0 = all.
    pub lift_mask: u32,
    /// The arguments that read the element a missing `else` of an `IF` gave as the text `FALSE`
    /// and as an empty text (a bit per position); the others read it as the number 0
    /// (`marker_table.rs`).
    pub marker_text: u32,
    pub marker_empty: u32,
    /// The arguments (a bit per position) whose boolean format the result of the function keeps
    /// when it is a number: `VALUE(TRUE())` is the number 1 that its cell shows as `TRUE`, where
    /// `ABS(TRUE())` is the number 1 (`keep_format_table.rs`). [`u32::MAX`] on the functions that
    /// give what one of their arguments gave (`IF`, `CHOOSE`…), which set the format themselves.
    pub keep_format: u32,
    /// The arguments that must be a range or an array and the error of a plain value there (see
    /// [`MATRIX_ARGS`]); empty for most functions.
    pub matrix: &'static [(u32, ErrorKind)],
    /// What `LibreOffice` does with each number of arguments (`arity.rs`); empty if not probed.
    pub lo_arity: &'static str,
    /// `LibreOffice` refuses some count that the engine takes (`AVERAGEIFS` with an even count).
    pub lo_irregular: bool,
    /// Takes the arrays of an operator over a range with their rows past the data not stored
    /// (`SUMPRODUCT`: they are counted, not read); any other function gets them stored.
    pub ghost_ok: bool,
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
            force_array: 0,
            range_error: 0,
            propagates: false,
            lift: false,
            lift_mask: 0,
            marker_text: 0,
            marker_empty: 0,
            keep_format: 0,
            matrix: &[],
            lo_arity: "",
            lo_irregular: false,
            ghost_ok: false,
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

    /// Every argument is evaluated in array context (the starting point; see [`NOT_FORCED`]).
    pub(crate) fn array_args(&mut self) -> &mut Self {
        self.force_array = u32::MAX;
        self
    }

    pub(crate) fn dynamic(&mut self) -> &mut Self {
        self.dynamic = true;
        self
    }
}

#[must_use]
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
    threed::register(&mut r);
    literal_error::register(&mut r);
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
        if NOT_FORCED.contains(&d.name) {
            d.force_array = 0;
        } else if let Some(&(_, mask)) = FORCE_ARRAY.iter().find(|(n, _)| *n == d.name) {
            d.force_array = mask;
        }
        if let Some(&(_, mask)) = RANGE_IS_ERROR.iter().find(|(n, _)| *n == d.name) {
            d.range_error = mask;
        }
        d.ghost_ok = d.name == "SUMPRODUCT";
        d.propagates = literal_error::propagates(d.name, d.ns);
        // The arguments that take an array written in the formula one element at a time (and a
        // range in array evaluation): what `LibreOffice` does with each function, probed.
        if let Ok(i) =
            lift_table::LIFT.binary_search_by(|&(n, ns, _)| (n, ns).cmp(&(d.name, d.ns.as_str())))
        {
            // What takes arrays whole is not lifted: the arguments that must be a range or an array
            // ([`MATRIX_ARGS`]) and the functions that make an array of the arrays they are given.
            let whole = MATRIX_ARGS
                .iter()
                .find(|(n, _)| *n == d.name)
                .map_or(0, |(_, m)| m.iter().fold(0, |acc, (mask, _)| acc | mask));
            let mask = if WHOLE_ARRAY.contains(&d.name) {
                0
            } else {
                lift_table::LIFT[i].2 & !whole
            };
            if mask != 0 {
                d.lift = true;
                d.lift_mask = mask;
            }
        }
        let key = (d.name, d.ns.as_str());
        let find = |t: &[(&str, &str, u32)]| {
            t.binary_search_by(|&(n, ns, _)| (n, ns).cmp(&key))
                .map_or(0, |i| t[i].2)
        };
        d.marker_text = find(marker_table::TEXT);
        d.marker_empty = find(marker_table::EMPTY);
        d.keep_format = if matches!(
            d.name,
            "IF" | "IFS" | "IFERROR" | "IFNA" | "CHOOSE" | "SWITCH"
        ) {
            u32::MAX
        } else {
            find(keep_format_table::KEEPS)
        };
        if let Some(&(_, matrix)) = MATRIX_ARGS.iter().find(|(n, _)| *n == d.name) {
            d.matrix = matrix;
        }
        if let Ok(i) =
            arity::ARITY.binary_search_by(|&(n, ns, _)| (n, ns).cmp(&(d.name, d.ns.as_str())))
        {
            d.lo_arity = arity::ARITY[i].2;
            let top = (d.max as usize).min(d.lo_arity.len() + 1);
            d.lo_irregular = (d.min as usize..=top).any(|n| lo_arity_refusal(d, n).is_some());
        }
    }
    let map =
        r.0.iter()
            .enumerate()
            .map(|(i, d)| ((d.ns, d.name), i as u16))
            .collect();
    Registry(r.0, map)
});

/// The function `name` (uppercase) of exactly this namespace.
#[must_use]
pub fn lookup_in(ns: Ns, name: &str) -> Option<FuncId> {
    REGISTRY.1.get(&(ns, name)).map(|&i| FuncId(i))
}

/// Resolve `name` (uppercase) written in `ns`: that namespace first, then `core`, then any other
/// (a formula that leaves out a prefix the function needs still finds it). `None`: unknown.
#[must_use]
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
#[must_use]
pub fn xlsx_bare(name: &str) -> Option<Ns> {
    namespaces::XLSX_BARE
        .iter()
        .find(|(n, _)| *n == name)
        .map(|&(_, ns)| ns)
}

/// Whether `LibreOffice` has a function of this name (uppercase, without prefix): see
/// `namespaces::CATALOGUE`. A function it has and the engine does not implement stops the build.
#[must_use]
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
#[must_use]
pub fn round_decimal(x: f64, digits: i32) -> f64 {
    math::round_dec(x, digits, math::Mode::Nearest)
}
