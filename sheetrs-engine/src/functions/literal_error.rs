//! An error written in the formula (`#NAME?`, `#NUM!`, `#DIV/0!`, `#N/A`, `#NULL!`, `#VALUE!`; a bare
//! `#REF!` in an XLSX is a name that does not exist, not an error) is not the error a calculation gave.
//!
//! Where a function wants a reference (`OFFSET(#NAME?,0,0)`, `SUMIF(#NAME?,1)`, `ROW(#NAME?)`) the token
//! is no reference and `LibreOffice` answers `Err:504` (parameter list), while the same call with a
//! calculated error (`OFFSET(NA(),0,0)`) is that error; where a function wants a value both are the
//! error (`OFFSET(Data!A1,#NAME?,0)` is `#NAME?`). `scripts/probe_literal_error.py` asks every function
//! of the registry and writes the table `LITERAL_ERROR` (`literal_error_table.rs`).
//!
//! An error a calculation gave in any argument is the answer, before the arguments are checked
//! (`SUMIF(NA(),1)` is `#N/A`, and so is `ACCRINT("a","a","a","a","a","a",NA())`, not the `#VALUE!` of the
//! text): [`propagated_error`], for the functions of
//! `PROPAGATING` (`propagating_error_table.rs`) (`scripts/probe_literal_error.py`).
//!
//! And a calculated error wins over a written one, whatever the order (`SUM(#NAME?,NA())` is `#N/A`):
//! `scripts/probe_literal_error.py` asks it of every function and writes
//! `COMPUTED_FIRST` (`computed_error_table.rs`), the functions that do not catch
//! the calculated error (`IFERROR(1/0,#NAME?)` is `#NAME?`).
//!
//! A call that has a written error where it wants a reference, or one in a function of that table next
//! to a calculated one, is not the function's answer: the compiler lowers it to `_LITERR`
//! (`sheetrs-compiler`, `error_where_a_reference_is_wanted`), which gives, in the order `LibreOffice` does
//! (checked with calls that have several errors):
//!
//! 1. the first error a calculation gave in the arguments;
//! 2. else, where the function wants a reference, the first error written where it wants a value
//!    (`OFFSET(#NUM!,#NAME?,0)` is `#NAME?`: the reference position loses to the value position),
//!    and `Err:504` when there is none;
//! 3. else what the function does.

use super::computed_error_table::COMPUTED_FIRST;
use super::literal_error_table::LITERAL_ERROR;
use super::propagating_error_table::PROPAGATING;
use super::{Reg, VAR};
use crate::engine::Res;
use crate::ir::Ns;
use crate::value::{ErrorKind, Value};

/// Names that are one function spelled two ways (`SKEW.P` is `SKEWP` in an ODS, `CHISQ.TEST` is `CHITEST`): the
/// tables are probed on the spelling an XLSX has, and the other spelling answers as it does.
const ALIASES: &[&[&str]] = &[
    &["B", "BINOM.DIST.RANGE"],
    &["BETA.INV", "BETAINV"],
    &["BINOM.DIST", "BINOMDIST"],
    &["BINOM.INV", "CRITBINOM"],
    &["CEILING.PRECISE", "ISO.CEILING"],
    &["CHISQ.DIST.RT", "CHIDIST"],
    &["CHISQ.INV.RT", "CHIINV"],
    &["CHISQ.TEST", "CHITEST"],
    &["CONFIDENCE", "CONFIDENCE.NORM"],
    &["CORREL", "PEARSON"],
    &["F.DIST", "FDIST"],
    &["F.INV", "FINV"],
    &["F.TEST", "FTEST"],
    &["FORECAST", "FORECAST.LINEAR"],
    &["GAMMA.DIST", "GAMMADIST"],
    &["GAMMA.INV", "GAMMAINV"],
    &["GAMMALN", "GAMMALN.PRECISE"],
    &["LOGINV", "LOGNORM.INV"],
    &["NORM.DIST", "NORMDIST"],
    &["NORM.INV", "NORMINV"],
    &["PERCENTILE", "PERCENTILE.INC"],
    &["PERCENTRANK", "PERCENTRANK.INC"],
    &["QUARTILE", "QUARTILE.INC"],
    &["SKEW.P", "SKEWP"],
    &["STDEV.P", "STDEVP"],
    &["T.INV.2T", "TINV"],
    &["T.TEST", "TTEST"],
    &["VAR.P", "VARP"],
    &["WEIBULL", "WEIBULL.DIST"],
    &["Z.TEST", "ZTEST"],
];

const NAMESPACES: [Ns; 5] = [
    Ns::Core,
    Ns::Microsoft,
    Ns::OpenOffice,
    Ns::LibreOffice,
    Ns::Legacy,
];

/// `find` on the function `name` of the namespace `ns`, else on the other spelling of it.
fn find_spelled<T>(name: &str, ns: Ns, find: impl Fn(&str, &str) -> Option<T>) -> Option<T> {
    find(name, ns.as_str()).or_else(|| {
        ALIASES
            .iter()
            .find(|group| group.contains(&name))?
            .iter()
            .filter(|peer| **peer != name)
            .find_map(|peer| NAMESPACES.iter().find_map(|n| find(peer, n.as_str())))
    })
}

/// The answer of the function `name` of the namespace `ns` to an error written as its argument number `arg`
/// (from 0), when it wants a reference or an array there and the token is neither: `Err:504` for `OFFSET`'s
/// reference, `Err:502` for the arrays of `CORREL`… `None` where it takes a value (the error written there
/// is that error).
///
/// `variadic` is whether it takes any number of arguments: then the arguments after the last one the
/// table lists are as the last character says, except in the functions whose arguments come in pairs
/// (`COUNTIFS(range, criterion, range, criterion…)`, and the `…IFS` that start with a sum range).
/// `count` is the number of arguments of the call: `LINEST(#NAME?)` is `Err:504` and `LINEST(#NAME?,Data!A1:A3)`
/// is `Err:502`.
#[must_use]
pub fn literal_answer(
    name: &str,
    ns: Ns,
    arg: usize,
    variadic: bool,
    count: usize,
) -> Option<ErrorKind> {
    let i = find_spelled(name, ns, |n, s| {
        LITERAL_ERROR
            .binary_search_by(|&(m, t, _)| (m, t).cmp(&(n, s)))
            .ok()
    })?;
    if variadic {
        let pair = match name {
            "COUNTIFS" => Some(arg.is_multiple_of(2)),
            "SUMIFS" | "AVERAGEIFS" | "MAXIFS" | "MINIFS" => Some(!arg.is_multiple_of(2)),
            "SORTBY" => Some(arg == 0 || !arg.is_multiple_of(2)),
            _ => None,
        };
        if let Some(wanted) = pair {
            return wanted.then_some(ErrorKind::Libre(504));
        }
    }
    if matches!(name, "LINEST" | "TREND") && arg == 0 {
        return Some(ErrorKind::Libre(if count == 1 { 504 } else { 502 }));
    }
    let positions = LITERAL_ERROR[i].2.as_bytes();
    let code = match positions.get(arg) {
        Some(c) => *c,
        None if variadic => *positions.last()?,
        None => return None,
    };
    match code {
        b'4' => Some(ErrorKind::Libre(504)),
        b'2' => Some(ErrorKind::Libre(502)),
        b'v' => Some(ErrorKind::Value),
        b'7' => Some(ErrorKind::Libre(517)),
        _ => None,
    }
}

/// The functions whose arguments are arrays (or lists of arrays) and that read an error written in the formula
/// there as an array of one element that is that error, not as an error to pass on: what each does with it
/// is what it does with an element that is an error (`SUMPRODUCT(1,#NAME?)` is `#NAME?`, `SUMPRODUCT("a",#NAME?)`
/// is 0, `SUMPRODUCT(#NAME?,Data!A1:A3)` is `#VALUE!`: the sizes are checked first).
/// `scripts/probe_literal_error.py --array` asks them.
const TOKEN_AS_ARRAY: &[&str] = &[
    "COVAR",
    "COVARIANCE.P",
    "HSTACK",
    "SUMPRODUCT",
    "SUMX2MY2",
    "SUMX2PY2",
    "SUMXMY2",
    "VSTACK",
];

/// The positions of the two series (values and timeline) of the functions of `FORECAST.ETS`: an error written in
/// one of them next to a series of several cells is `Err:502`, as a value that is no series is. Written next to a
/// plain value, or in the arguments that follow, `LibreOffice` answers a number that comes from reading the error as
/// a number (`FORECAST.ETS(7,…,…,#NAME?)` is 9.22337203685478E+18, the largest integer): see `FIDELITY.md`.
#[must_use]
pub fn ets_series(name: &str) -> Option<(usize, usize)> {
    match name {
        "FORECAST.ETS"
        | "FORECAST.ETS.ADD"
        | "FORECAST.ETS.MULT"
        | "FORECAST.ETS.CONFINT"
        | "FORECAST.ETS.PI.ADD"
        | "FORECAST.ETS.PI.MULT" => Some((1, 2)),
        "FORECAST.ETS.SEASONALITY"
        | "FORECAST.ETS.STAT"
        | "FORECAST.ETS.STAT.ADD"
        | "FORECAST.ETS.STAT.MULT" => Some((0, 1)),
        _ => None,
    }
}

/// Whether the argument number `arg` (from 0) of the function `name` is a criterion: one value, the criteria
/// of the `…IFS` come after each range (`SUMIFS(sum, range1, criterion1, …)`; `COUNTIFS(range1, criterion1, …)`).
#[must_use]
pub fn is_criterion(name: &str, arg: usize) -> bool {
    match name {
        "SUMIFS" | "AVERAGEIFS" | "MAXIFS" | "MINIFS" => arg >= 2 && arg.is_multiple_of(2),
        "COUNTIFS" => !arg.is_multiple_of(2),
        _ => false,
    }
}

/// Whether an error written as an argument of the function `name` is an array of one element that is that
/// error (the list `TOKEN_AS_ARRAY` of this module).
#[must_use]
pub fn token_as_array(name: &str, _ns: Ns) -> bool {
    TOKEN_AS_ARRAY.contains(&name)
}

/// Whether the function wants a reference or an array as its argument number `arg`: see [`literal_answer`].
#[must_use]
pub fn wants_reference(name: &str, ns: Ns, arg: usize, variadic: bool) -> bool {
    literal_answer(name, ns, arg, variadic, usize::MAX).is_some()
}

/// Whether an error a calculation gave wins over an error written in the formula in a call of the function
/// `name` of the namespace `ns` (`SUM(#NAME?,NA())` is `#N/A`): the functions of
/// `COMPUTED_FIRST` (`computed_error_table.rs`).
#[must_use]
pub fn computed_error_first(name: &str, ns: Ns) -> bool {
    find_spelled(name, ns, |n, s| {
        COMPUTED_FIRST
            .binary_search_by(|&(m, t)| (m, t).cmp(&(n, s)))
            .ok()
    })
    .is_some()
}

/// Whether the function `name` of the namespace `ns` answers an error a calculation gave in any of its
/// arguments, whatever the other arguments are.
#[must_use]
pub fn propagates(name: &str, ns: Ns) -> bool {
    find_spelled(name, ns, |n, s| {
        PROPAGATING
            .binary_search_by(|&(m, t)| (m, t).cmp(&(n, s)))
            .ok()
    })
    .is_some()
}

/// The error a call answers when an argument is an error a calculation gave: the first one, before the
/// arguments are checked (`SUMIF(NA(),1)` is `#N/A`, `ACCRINT("a",…,NA())` is `#N/A`, not the `#VALUE!` of the
/// text), for the functions that do (`FuncDef::propagates`).
#[must_use]
pub fn propagated_error(def: &super::FuncDef, args: &[Value]) -> Option<ErrorKind> {
    if !def.propagates {
        return None;
    }
    args.iter().find_map(|v| match v {
        Value::Error(kind) => Some(*kind),
        _ => None,
    })
}

/// `_LITERR(answer, calculated…)`: the first error among `calculated` (the arguments of a call that can be an
/// error a calculation gave), and else `answer` (an error written in the call, or the value of the call).
pub(super) fn register(r: &mut Reg) {
    r.eager("_LITERR", 1, VAR, |e, a| -> Res<Value> {
        // A range among them is one value here (a criterion): met by the cell of the formula, `#VALUE!` where
        // it has none.
        if let Some(kind) = a[1..].iter().find_map(|v| match v {
            Value::Error(kind) => Some(*kind),
            Value::Ref(_) => match e.scalar(v.clone()) {
                Value::Error(kind) => Some(kind),
                _ => None,
            },
            _ => None,
        }) {
            return Err(kind);
        }
        match &a[0] {
            Value::Error(kind) => Err(*kind),
            v => Ok(v.clone()),
        }
    });
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_table_is_sorted_by_function_and_namespace() {
        assert!(
            LITERAL_ERROR
                .windows(2)
                .all(|w| (w[0].0, w[0].1) < (w[1].0, w[1].1))
        );
    }

    #[test]
    fn the_functions_where_a_calculated_error_wins_are_sorted_and_have_the_usual_ones() {
        assert!(COMPUTED_FIRST.windows(2).all(|w| w[0] < w[1]));
        assert!(computed_error_first("SUM", Ns::Core));
        assert!(computed_error_first("ROUND", Ns::Core));
        // a function that catches the calculated error is not there
        assert!(!computed_error_first("IFERROR", Ns::Core));
    }

    #[test]
    fn a_function_spelled_with_or_without_the_dot_wants_a_reference_and_propagates_errors_alike() {
        assert!(wants_reference("SKEWP", Ns::Core, 0, true));
        assert!(wants_reference("SKEW.P", Ns::Core, 0, true));
        assert!(wants_reference("CHITEST", Ns::Core, 0, false));
        assert!(propagates("SKEWP", Ns::Core));
    }

    #[test]
    fn a_function_wants_a_reference_where_the_table_says() {
        assert!(wants_reference("OFFSET", Ns::Core, 0, false));
        assert!(!wants_reference("OFFSET", Ns::Core, 1, false));
        assert!(wants_reference("SUMIF", Ns::Core, 2, false));
        assert!(!wants_reference("SUMIF", Ns::Core, 1, false));
        assert!(!wants_reference("SUM", Ns::Core, 0, true));
        // the rest of the arguments of a function with any number of them
        assert!(wants_reference("AVEDEV", Ns::Core, 7, true));
        assert!(!wants_reference("NPV", Ns::Core, 0, true));
        assert!(wants_reference("NPV", Ns::Core, 9, true));
        // arguments that come in pairs
        assert!(wants_reference("COUNTIFS", Ns::Core, 4, true));
        assert!(!wants_reference("COUNTIFS", Ns::Core, 5, true));
        assert!(wants_reference("SUMIFS", Ns::Core, 5, true));
        assert!(!wants_reference("SUMIFS", Ns::Core, 4, true));
    }
}
