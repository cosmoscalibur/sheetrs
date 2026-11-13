//! Logical functions, special forms and information functions.

use super::{Reg, VAR, num};
use crate::coerce;
use crate::engine::{Engine, Res};
use crate::ir::Expr;
use crate::value::compare;
use crate::value::{ErrorKind, Value};
use std::cmp::Ordering;
use std::sync::Arc;

/// The elements of a value that is an array: an array, or a range of several cells in array
/// evaluation; `None` for a single value.
fn elements(e: &mut Engine, v: &Value) -> Option<Arc<crate::value::Array>> {
    match v {
        Value::Array(m) if m.ghost.is_some() => Some(Arc::new(m.expanded())),
        Value::Array(m) => Some(Arc::clone(m)),
        Value::Ref(r) if !r.is_single() && e.in_array_context() => {
            let m = e.range_array_full(*r);
            Some(if m.ghost.is_some() {
                Arc::new(m.expanded())
            } else {
                m
            })
        }
        _ => None,
    }
}

/// The element `(i, j)` of what a function chooses from: an array is read at the same place (a row
/// or a column of one is spread, a place it does not have is `#N/A`), a single value is the same for
/// every element.
fn pick(m: Option<&crate::value::Array>, single: &Value, i: u32, j: u32) -> Value {
    // The elements of an array are numbers: what a branch gives that is a boolean is 1 or 0. The
    // element that a missing `else` gave, chosen again by another `IF`, is an empty text that
    // `COUNTA` counts.
    let number = |v: Value| match v {
        Value::Bool(b) => Value::Number(f64::from(u8::from(b))),
        v if v.is_empty_path() => Value::text(""),
        v => v,
    };
    match m {
        None => number(single.clone()),
        Some(m) => {
            let (r, c) = (
                if m.rows == 1 { 0 } else { i },
                if m.cols == 1 { 0 } else { j },
            );
            if r < m.rows && c < m.cols {
                number(m.get(r, c).clone())
            } else {
                Value::Error(ErrorKind::Na)
            }
        }
    }
}

/// What a branch of `IF` or `CHOOSE` leaves as the format of the cell, in `LibreOffice`: a
/// boolean (`TRUE()`, `NOT(…)`, a comparison, a branch that is itself such an array) makes the
/// cell show booleans, a number from a function or an operator (`SUM(1)`, `5+1`, `DATE(…)`) shows
/// numbers, and a constant, a reference or a text leaves what the last branch before it set.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Format {
    Logical,
    Number,
    Keeps,
}

/// The format a branch sets, from what it is and what it gave.
fn format_of(x: &Expr, v: &Value) -> Format {
    match x {
        Expr::Const(Value::Bool(_)) => Format::Logical,
        Expr::Const(_) | Expr::Cell(_) | Expr::Range(_) | Expr::Missing => Format::Keeps,
        Expr::Unary(_, inner) if matches!(**inner, Expr::Const(_)) => Format::Keeps,
        _ => match v {
            Value::Bool(_) => Format::Logical,
            Value::Array(m) if m.logical => Format::Logical,
            Value::Number(_) | Value::Array(_) => Format::Number,
            _ => Format::Keeps,
        },
    }
}

/// What a branch of `IF` or `CHOOSE` is for an array of conditions: the value, with its elements
/// when it is an array, and the format it sets.
fn branch(e: &mut Engine, x: &Expr) -> (Value, Option<Arc<crate::value::Array>>, Format) {
    let v = match x {
        Expr::Missing => Value::Number(0.0),
        x => e.eval(x),
    };
    let format = format_of(x, &v);
    let m = elements(e, &v);
    let single = if m.is_some() { v } else { e.scalar(v) };
    (single, m, format)
}

/// `IF` over an array of conditions: the array of what each one chooses. The cell that holds it
/// shows booleans when the last branch that ran, element by element, set a boolean format.
fn if_matrix(e: &mut Engine, cond: &crate::value::Array, a: &[Expr]) -> Value {
    let (then_v, then_m, then_f) = branch(e, &a[1]);
    let other = a.get(2).map(|x| branch(e, x));
    let mut format = Format::Keeps;
    let mut data = Vec::with_capacity(cond.data.len());
    for i in 0..cond.rows {
        for j in 0..cond.cols {
            let x = cond.get(i, j);
            data.push(
                match coerce::to_bool(&Engine::element_of_a_call(x), &e.opts.locale) {
                    Err(k) => Value::Error(k),
                    Ok(true) => {
                        format = if then_f == Format::Keeps {
                            format
                        } else {
                            then_f
                        };
                        pick(then_m.as_deref(), &then_v, i, j)
                    }
                    Ok(false) => match &other {
                        Some((v, m, f)) => {
                            format = if *f == Format::Keeps { format } else { *f };
                            pick(m.as_deref(), v, i, j)
                        }
                        None => Value::empty_path(),
                    },
                },
            );
        }
    }
    Value::Array(Arc::new(
        crate::value::Array::new(cond.rows, cond.cols, data)
            .with_logical(format == Format::Logical)
            .with_jump_false(),
    ))
}

/// `CHOOSE` over an array of indexes (the format of the cell as in `if_matrix`).
fn choose_matrix(e: &mut Engine, index: &crate::value::Array, a: &[Expr]) -> Value {
    let branches: Vec<_> = a[1..].iter().map(|x| branch(e, x)).collect();
    let mut format = Format::Keeps;
    let mut data = Vec::with_capacity(index.data.len());
    for i in 0..index.rows {
        for j in 0..index.cols {
            let x = index.get(i, j).clone();
            data.push(match e.num(&x) {
                Err(k) => Value::Error(k),
                Ok(n) => {
                    let n = n.trunc();
                    if n < 1.0 || n as usize > branches.len() {
                        Value::Error(ErrorKind::Value)
                    } else {
                        let (v, m, f) = &branches[n as usize - 1];
                        if *f != Format::Keeps {
                            format = *f;
                        }
                        pick(m.as_deref(), v, i, j)
                    }
                }
            });
        }
    }
    Value::Array(Arc::new(
        crate::value::Array::new(index.rows, index.cols, data)
            .with_logical(format == Format::Logical)
            .with_jump_false(),
    ))
}

/// `IF(condition)`: the condition as a boolean. An array of conditions is read element by element, a
/// true one being the number 1 and a false one `FALSE`, as `LibreOffice` has them.
fn if_alone(e: &mut Engine, c: &Value) -> Res<Value> {
    if let Value::Array(m) = c {
        let data = m
            .data
            .iter()
            .map(
                |x| match coerce::to_bool(&Engine::element_of_a_call(x), &e.opts.locale) {
                    Ok(true) => Value::Number(1.0),
                    Ok(false) => Value::Bool(false),
                    Err(k) => Value::Error(k),
                },
            )
            .collect();
        return Ok(Value::Array(std::sync::Arc::new(crate::value::Array::new(
            m.rows, m.cols, data,
        ))));
    }
    Ok(Value::Bool(e.boolean(c)?))
}

pub(super) fn register(r: &mut Reg) {
    r.lazy("IF", 1, 3, |e, a| {
        let c = e.eval(&a[0]);
        if a.len() == 1 {
            return if_alone(e, &c);
        }
        if let Some(m) = elements(e, &c) {
            return Ok(if_matrix(e, &m, a));
        }
        let branch = if e.boolean(&c)? {
            Some(&a[1])
        } else {
            a.get(2)
        };
        match branch {
            None => Ok(Value::Bool(false)),
            Some(Expr::Missing) => Ok(Value::Number(0.0)),
            Some(x) => {
                e.logical_fmt = crate::engine::FMT_NONE;
                Ok(e.eval(x))
            }
        }
    });
    r.lazy("IFS", 0, VAR, |e, a| {
        // A condition with no value: `IFS(TRUE)` lacks a variable (`Err:511`), while the one that
        // ends a longer list is never looked at.
        if let [only] = a {
            let c = e.eval(only);
            return if e.boolean(&c)? {
                Err(ErrorKind::Libre(511))
            } else {
                Err(ErrorKind::Na)
            };
        }
        for p in a.chunks(2) {
            if p.len() < 2 {
                break;
            }
            let c = e.eval(&p[0]);
            if e.boolean(&c)? {
                e.logical_fmt = crate::engine::FMT_NONE;
                return Ok(e.eval(&p[1]));
            }
        }
        Err(ErrorKind::Na)
    });
    // `IFERROR(, x)` lacks the value (`Err:518`); `IFNA(#N/A, )` has nothing to give instead: the error.
    r.lazy("IFERROR", 2, 2, |e, a| {
        if matches!(a[0], Expr::Missing) {
            return Err(ErrorKind::Libre(518));
        }
        guard(e, a, |_| true)
    });
    r.lazy("IFNA", 2, 2, |e, a| {
        if matches!(a[0], Expr::Missing) {
            return Err(ErrorKind::Libre(518));
        }
        if matches!(a[1], Expr::Missing) {
            let v = e.eval(&a[0]);
            let probe = e.scalar(v.clone());
            return match probe {
                Value::Error(ErrorKind::Na) => Err(ErrorKind::Na),
                _ => Ok(v),
            };
        }
        guard(e, a, |k| k == ErrorKind::Na)
    });
    r.lazy("CHOOSE", 1, VAR, |e, a| {
        let i = e.eval(&a[0]);
        if a.len() > 1
            && let Some(m) = elements(e, &i)
        {
            return Ok(choose_matrix(e, &m, a));
        }
        let i = e.num(&i)?.trunc();
        // No value to choose from: an invalid argument, whatever the index.
        if a.len() == 1 {
            return Err(ErrorKind::Libre(502));
        }
        // An index outside the values is an invalid argument.
        if i < 1.0 || i as usize >= a.len() {
            return Err(ErrorKind::Libre(502));
        }
        e.logical_fmt = crate::engine::FMT_NONE;
        Ok(e.eval(&a[i as usize]))
    });
    r.lazy("SWITCH", 3, VAR, |e, a| {
        let x = e.eval(&a[0]);
        let x = e.scalar(x);
        if let Value::Error(k) = x {
            return Err(k);
        }
        let mut rest = &a[1..];
        // With no match and no default the answer is the error of the last value to match, when it is one
        // (`SWITCH(1,NA(),5,1/0,6)` is `#DIV/0!`), else `#N/A`.
        let mut last_error = None;
        while rest.len() >= 2 {
            let c = e.eval(&rest[0]);
            let c = e.scalar(c);
            last_error = match &c {
                Value::Error(k) => Some(*k),
                _ => None,
            };
            if compare(&x, &c, e.collator(e.opts.search.case_sensitive)) == Ordering::Equal
                && std::mem::discriminant(&x) == std::mem::discriminant(&c)
            {
                e.logical_fmt = crate::engine::FMT_NONE;
                return Ok(e.eval(&rest[1]));
            }
            rest = &rest[2..];
        }
        match rest.first() {
            Some(d) => {
                e.logical_fmt = crate::engine::FMT_NONE;
                Ok(e.eval(d))
            }
            None => Err(last_error.unwrap_or(ErrorKind::Na)),
        }
    });
    r.eager("AND", 1, VAR, |e, a| fold_bool(e, a, true, |x, y| x && y));
    r.eager("OR", 1, VAR, |e, a| fold_bool(e, a, false, |x, y| x || y));
    r.eager("XOR", 1, VAR, |e, a| fold_bool(e, a, false, |x, y| x ^ y));
    r.eager("NOT", 1, 1, |e, a| Ok(Value::Bool(!e.boolean(&a[0])?)));
    r.eager("TRUE", 0, 0, |_, _| Ok(Value::Bool(true)));
    r.eager("FALSE", 0, 0, |_, _| Ok(Value::Bool(false)));

    r.eager("ISBLANK", 1, 1, |e, a| {
        Ok(Value::Bool(matches!(e.scalar(a[0].clone()), Value::Empty)))
    });
    // LibreOffice has no boolean type: TRUE() is a number. A formula that ended in a blank
    // holds 0 but is not a number.
    r.eager("ISNUMBER", 1, 1, |e, a| {
        if let Value::Ref(r) = &a[0]
            && r.is_single()
            && e.blank_result(r.top_left())
        {
            return Ok(Value::Bool(false));
        }
        Ok(Value::Bool(matches!(
            e.scalar(a[0].clone()),
            Value::Number(_) | Value::Bool(_)
        )))
    });
    r.eager("ISTEXT", 1, 1, |e, a| {
        Ok(Value::Bool(matches!(
            e.scalar(a[0].clone()),
            Value::Text(_)
        )))
    });
    r.eager("ISNONTEXT", 1, 1, |e, a| {
        Ok(Value::Bool(!matches!(
            e.scalar(a[0].clone()),
            Value::Text(_)
        )))
    });
    r.eager("ISLOGICAL", 1, 1, |e, a| {
        // An array holds numbers: what a comparison put in it is not a boolean any more (the array
        // written in the formula with `TRUE` first is, for the compiler: see `formula.rs`).
        if matches!(a[0], Value::Array(_)) {
            return Ok(Value::Bool(false));
        }
        // `-TRUE()` is a number that `LibreOffice` takes for a boolean here; `VALUE(TRUE())` is not.
        let negated = e.logical_fmt == crate::engine::FMT_NEGATED;
        Ok(Value::Bool(
            negated || matches!(e.scalar(a[0].clone()), Value::Bool(_)),
        ))
    });
    r.eager("ISERROR", 1, 1, |e, a| {
        Ok(Value::Bool(e.scalar(a[0].clone()).is_error()))
    });
    r.eager("ISERR", 1, 1, |e, a| {
        Ok(Value::Bool(
            matches!(e.scalar(a[0].clone()), Value::Error(k) if k != ErrorKind::Na),
        ))
    });
    r.eager("ISNA", 1, 1, |e, a| {
        Ok(Value::Bool(matches!(
            e.scalar(a[0].clone()),
            Value::Error(ErrorKind::Na)
        )))
    });
    r.eager("ISREF", 1, 1, |_, a| {
        Ok(Value::Bool(matches!(a[0], Value::Ref(_))))
    });
    r.eager("ISEVEN", 1, 1, |e, a| {
        Ok(Value::Bool(e.num(&a[0])?.trunc() % 2.0 == 0.0))
    });
    r.eager("ISODD", 1, 1, |e, a| {
        Ok(Value::Bool(e.num(&a[0])?.trunc().abs() % 2.0 == 1.0))
    });
    r.eager("NA", 0, 0, |_, _| Err(ErrorKind::Na));
    r.eager("N", 1, 1, |e, a| match e.scalar(a[0].clone()) {
        Value::Number(n) => Ok(Value::Number(n)),
        Value::Bool(b) => num(f64::from(u8::from(b))),
        Value::Error(k) => Err(k),
        _ => num(0.0),
    });
    // The internal codes of LibreOffice (`Err:532` is #DIV/0!), not Excel's 1 to 7.
    r.eager("ERRORTYPE", 1, 1, |e, a| match e.scalar(a[0].clone()) {
        Value::Error(k) => match k {
            ErrorKind::Null => num(521.0),
            ErrorKind::DivZero => num(532.0),
            ErrorKind::Ref => num(524.0),
            ErrorKind::Name => num(525.0),
            ErrorKind::Num => num(503.0),
            ErrorKind::Na => num(32767.0),
            _ => {
                crate::notes::note(crate::notes::ERRORTYPE_VALUE);
                Err(ErrorKind::Unsupported)
            }
        },
        _ => Err(ErrorKind::Na),
    });
    // The errors only `LibreOffice` has (`Err:502`, `Err:504`…) have no number: `#N/A`.
    r.eager("ERROR.TYPE", 1, 1, |e, a| match e.scalar(a[0].clone()) {
        Value::Error(ErrorKind::Libre(_)) => Err(ErrorKind::Na),
        Value::Error(k) => num(k.code()),
        _ => Err(ErrorKind::Na),
    });
}

/// `IFERROR`/`IFNA`: evaluate the first argument, fall back to the second when
/// it is an error accepted by `catches`.
#[allow(clippy::unnecessary_wraps)] // same fallible signature as its siblings, which share call sites
fn guard(e: &mut Engine, a: &[Expr], catches: fn(ErrorKind) -> bool) -> Res<Value> {
    let v = e.eval(&a[0]);
    // Outside array evaluation a range is met by the formula cell, as in `LibreOffice`: what it
    // meets is the value, and when they do not meet that is the `#VALUE!` the guard catches.
    let v = match v {
        Value::Ref(r) if !r.is_single() && !e.in_array_context() => e.scalar(v),
        v => v,
    };
    // An array: each element that is an error this catches is the alternative.
    if let Value::Array(m) = &v {
        if !m
            .data
            .iter()
            .any(|x| matches!(x, Value::Error(k) if catches(*k)))
        {
            // The guard gives the numbers back: what `IF` set as a boolean format is not kept.
            // It is what a jump function (`IF`, `CHOOSE`, `IFERROR`) gave: a function that takes
            // one value gets its first element outside array evaluation.
            return Ok(Value::Array(Arc::new(
                (**m).clone().with_logical(false).with_jump_false(),
            )));
        }
        let (alt, alt_m, _) = branch(e, &a[1]);
        let mut data = Vec::with_capacity(m.data.len());
        for i in 0..m.rows {
            for j in 0..m.cols {
                data.push(match m.get(i, j) {
                    Value::Error(k) if catches(*k) => pick(alt_m.as_deref(), &alt, i, j),
                    x => x.clone(),
                });
            }
        }
        return Ok(Value::Array(Arc::new(
            crate::value::Array::new(m.rows, m.cols, data).with_jump_false(),
        )));
    }
    let probe = match &v {
        Value::Ref(r) if r.is_single() => e.scalar(v.clone()),
        Value::Error(_) => v.clone(),
        _ => return Ok(v),
    };
    match probe {
        Value::Error(k) if catches(k) => Ok(match &a[1] {
            Expr::Missing => Value::Number(0.0),
            x => {
                e.logical_fmt = crate::engine::FMT_NONE;
                e.eval(x)
            }
        }),
        _ => Ok(v),
    }
}

fn fold_bool(e: &mut Engine, a: &[Value], init: bool, f: fn(bool, bool) -> bool) -> Res<Value> {
    let mut acc = init;
    let mut seen = false;
    for v in a {
        match v {
            Value::Ref(_) | Value::Array(_) => {
                for x in &e.array_of(v.clone()).data {
                    match x {
                        Value::Bool(b) => {
                            acc = f(acc, *b);
                            seen = true;
                        }
                        Value::Number(n) => {
                            acc = f(acc, *n != 0.0);
                            seen = true;
                        }
                        Value::Error(k) => return Err(*k),
                        // An array written in the formula, or one a function made, with a text in
                        // it is an invalid argument; the texts of a range are left out.
                        Value::Text(_) if matches!(v, Value::Array(_)) => {
                            return Err(ErrorKind::Libre(502));
                        }
                        _ => {}
                    }
                }
            }
            v => {
                acc = f(acc, coerce::to_bool(v, &e.opts.locale)?);
                seen = true;
            }
        }
    }
    if seen {
        Ok(Value::Bool(acc))
    } else {
        Err(ErrorKind::Value)
    }
}
