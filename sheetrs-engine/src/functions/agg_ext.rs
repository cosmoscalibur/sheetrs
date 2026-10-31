//! `SUBTOTAL` and `AGGREGATE`: aggregates that can skip hidden rows, errors and
//! the subtotals inside their own range.

use super::stats::variance;
use super::{Reg, VAR, num};
use crate::addr::CellRef;
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

struct Skip {
    hidden: bool,
    errors: bool,
    nested: bool,
}

/// The cells of the references and arrays, in order, minus what `skip` leaves out.
fn cells(e: &mut Engine, args: &[Value], skip: &Skip) -> Res<Vec<Value>> {
    let mut out = Vec::new();
    for a in args {
        match a {
            Value::Ref(r) => {
                for row in r.r0..=r.r1 {
                    if skip.hidden && e.is_row_hidden(r.sheet, row) {
                        continue;
                    }
                    for col in r.c0..=r.c1 {
                        let at = CellRef::new(r.sheet, row, col);
                        if skip.nested && e.is_subtotal_cell(at) {
                            continue;
                        }
                        out.push(e.cell_value(at));
                    }
                }
            }
            Value::Array(m) => out.extend(m.data.iter().cloned()),
            v => out.push(v.clone()),
        }
    }
    if skip.errors {
        out.retain(|v| !v.is_error());
    }
    Ok(out)
}

fn numbers_of(v: &[Value]) -> Res<Vec<f64>> {
    let mut out = Vec::new();
    for x in v {
        match x {
            Value::Number(n) => out.push(*n),
            Value::Error(k) => return Err(*k),
            _ => {}
        }
    }
    Ok(out)
}

fn mean(v: &[f64]) -> f64 {
    v.iter().sum::<f64>() / v.len() as f64
}

fn dev(v: &[f64], sample: bool) -> Res<f64> {
    let n = v.len() as f64;
    if n < if sample { 2.0 } else { 1.0 } {
        return Err(ErrorKind::DivZero);
    }
    let m = mean(v);
    Ok(v.iter().map(|x| (x - m).powi(2)).sum::<f64>() / if sample { n - 1.0 } else { n })
}

/// Functions 1–11 shared by both: AVERAGE, COUNT, COUNTA, MAX, MIN, PRODUCT,
/// STDEV, STDEVP, SUM, VAR, VARP.
fn simple(code: i64, v: &[Value]) -> Res<Value> {
    if code == 3 {
        return num(v.iter().filter(|x| !matches!(x, Value::Empty)).count() as f64);
    }
    let n = numbers_of(v)?;
    match code {
        1 => {
            if n.is_empty() {
                Err(ErrorKind::DivZero)
            } else {
                num(mean(&n))
            }
        }
        2 => num(n.len() as f64),
        4 => num(n.iter().copied().reduce(f64::max).unwrap_or(0.0)),
        5 => num(n.iter().copied().reduce(f64::min).unwrap_or(0.0)),
        6 => num(if n.is_empty() {
            0.0
        } else {
            n.iter().product()
        }),
        7 => dev(&n, true).map(f64::sqrt).and_then(num),
        8 => dev(&n, false).map(f64::sqrt).and_then(num),
        9 => num(n.iter().sum()),
        10 => dev(&n, true).and_then(num),
        11 => dev(&n, false).and_then(num),
        _ => Err(ErrorKind::Value),
    }
}

pub(super) fn register(r: &mut Reg) {
    r.eager("SUBTOTAL", 2, VAR, |e, a| {
        let code = e.num(&a[0])?.trunc() as i64;
        let hidden = code > 100;
        let base = if hidden { code - 100 } else { code };
        if !(1..=11).contains(&base) {
            return Err(ErrorKind::Value);
        }
        let v = cells(
            e,
            &a[1..],
            &Skip {
                hidden,
                errors: false,
                nested: true,
            },
        )?;
        simple(base, &v)
    });
    r.eager("AGGREGATE", 3, VAR, |e, a| {
        let code = e.num(&a[0])?.trunc() as i64;
        let opt = e.num(&a[1])?.trunc() as i64;
        if !(1..=19).contains(&code) || !(0..=7).contains(&opt) {
            return Err(ErrorKind::Value);
        }
        let skip = Skip {
            hidden: matches!(opt, 1 | 3 | 5 | 7),
            errors: matches!(opt, 2 | 3 | 6 | 7),
            nested: opt <= 3,
        };
        if code <= 13 {
            let v = cells(e, &a[2..], &skip)?;
            return match code {
                12 => {
                    let mut n = numbers_of(&v)?;
                    if n.is_empty() {
                        return Err(ErrorKind::Value);
                    }
                    n.sort_by(f64::total_cmp);
                    let m = n.len();
                    num(if m % 2 == 1 {
                        n[m / 2]
                    } else {
                        (n[m / 2 - 1] + n[m / 2]) / 2.0
                    })
                }
                13 => {
                    let n = numbers_of(&v)?;
                    let mut best: Option<(usize, f64)> = None;
                    for x in &n {
                        let c = n.iter().filter(|y| *y == x).count();
                        if c > 1 && best.is_none_or(|(b, _)| c > b) {
                            best = Some((c, *x));
                        }
                    }
                    best.map_or(Err(ErrorKind::Na), |(_, x)| num(x))
                }
                c => simple(c, &v),
            };
        }
        // 14–19 take an array and the k or quantile.
        if a.len() < 4 {
            return Err(ErrorKind::Value);
        }
        let v = cells(e, &a[2..3], &skip)?;
        let mut n = numbers_of(&v)?;
        n.sort_by(f64::total_cmp);
        let k = e.num(&a[3])?;
        let len = n.len() as f64;
        match code {
            14 | 15 => {
                let i = k.trunc();
                if i < 1.0 || i > len {
                    return Err(ErrorKind::Domain);
                }
                let i = i as usize;
                num(if code == 14 { n[n.len() - i] } else { n[i - 1] })
            }
            16..=19 => {
                let q = if matches!(code, 17 | 19) {
                    k.trunc() / 4.0
                } else {
                    k
                };
                let excl = matches!(code, 18 | 19);
                if n.is_empty() {
                    return Err(ErrorKind::Domain);
                }
                if excl {
                    if q < 1.0 / (len + 1.0) || q > len / (len + 1.0) {
                        return Err(ErrorKind::Domain);
                    }
                    let pos = q * (len + 1.0);
                    let lo = pos.floor() as usize;
                    num(n[lo - 1] + (pos - lo as f64) * (n[lo.min(n.len() - 1)] - n[lo - 1]))
                } else {
                    if !(0.0..=1.0).contains(&q) {
                        return Err(ErrorKind::Domain);
                    }
                    let pos = q * (len - 1.0);
                    let lo = pos.floor() as usize;
                    let hi = (lo + 1).min(n.len() - 1);
                    num(n[lo] + (pos - lo as f64) * (n[hi] - n[lo]))
                }
            }
            _ => Err(ErrorKind::Value),
        }
    })
    .array_args();
    let _ = variance;
}
