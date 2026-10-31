//! `LINEST`, `LOGEST`, `TREND`, `GROWTH`, `FREQUENCY` and `MODE.MULT`: the functions that answer an
//! array computed from a least-squares fit or a count. What LibreOffice 26.8 answers for the
//! shapes and values outside the plain case is what `tests/assets/semantics.*` holds.

use super::{Reg, VAR, opt};
use crate::engine::{Engine, Res};
use crate::value::{Array, ErrorKind, Value};
use std::sync::Arc;

const DOMAIN: ErrorKind = ErrorKind::Domain;

fn array(rows: usize, cols: usize, data: Vec<Value>) -> Value {
    Value::Array(Arc::new(Array::new(rows as u32, cols as u32, data)))
}

fn na() -> Value {
    Value::Error(ErrorKind::Na)
}

/// A rectangle of numbers; a cell that is not one (text, blank, error) is outside the domain.
struct Matrix {
    rows: usize,
    cols: usize,
    data: Vec<f64>,
}

impl Matrix {
    fn at(&self, r: usize, c: usize) -> f64 {
        self.data[r * self.cols + c]
    }
}

fn real_matrix(e: &mut Engine, v: &Value) -> Res<Matrix> {
    if let Value::Error(k) = v {
        return Err(*k);
    }
    let m = e.array_of(v.clone());
    let mut data = Vec::with_capacity(m.data.len());
    for x in &m.data {
        data.push(match x {
            Value::Number(n) => *n,
            Value::Bool(b) => f64::from(u8::from(*b)),
            _ => return Err(DOMAIN),
        });
    }
    Ok(Matrix {
        rows: m.rows as usize,
        cols: m.cols as usize,
        data,
    })
}

/// A flag argument (`const`, `stats`): the default when omitted, a blank is false and text that is
/// not a boolean is false.
fn flag(e: &mut Engine, a: &[Value], i: usize, default: bool) -> Res<bool> {
    let Some(v) = opt(a, i) else {
        return Ok(default);
    };
    match e.scalar(v.clone()) {
        Value::Error(k) => Err(k),
        Value::Text(t) => {
            Ok(crate::coerce::to_bool(&Value::Text(t), &e.opts.locale).unwrap_or(false))
        }
        v => crate::coerce::to_bool(&v, &e.opts.locale),
    }
}

/// The regressors of a fit: `k` columns of `n` values, and whether the data lie along rows.
struct Design {
    y: Vec<f64>,
    x: Vec<Vec<f64>>,
    by_rows: bool,
    single: bool,
}

fn design(e: &mut Engine, ys: &Value, xs: Option<&Value>) -> Res<Design> {
    let y = real_matrix(e, ys)?;
    if y.rows > 1 && y.cols > 1 {
        return Err(DOMAIN);
    }
    let by_rows = y.rows == 1 && y.cols > 1;
    let n = y.data.len();
    let x = match xs {
        None => vec![(1..=n).map(|i| i as f64).collect()],
        Some(xs) => {
            let m = real_matrix(e, xs)?;
            if by_rows {
                if m.cols != n {
                    return Err(DOMAIN);
                }
                (0..m.rows)
                    .map(|r| (0..n).map(|c| m.at(r, c)).collect())
                    .collect()
            } else {
                if m.rows != n {
                    return Err(DOMAIN);
                }
                (0..m.cols)
                    .map(|c| (0..n).map(|r| m.at(r, c)).collect())
                    .collect()
            }
        }
    };
    let single = x.len() == 1;
    Ok(Design {
        y: y.data,
        x,
        by_rows,
        single,
    })
}

struct Fit {
    /// The coefficients in the order of the regressors, the intercept last when there is one.
    beta: Vec<f64>,
    /// The diagonal of `(X'X)^-1`.
    inv_diag: Vec<f64>,
    ssresid: f64,
    ssreg: f64,
    df: f64,
    constant: bool,
    /// The residual is rounding noise of a perfect fit.
    noise: bool,
}

/// Least squares by Householder reflections; a regressor that is a combination of the others
/// (or all zero) is `#VALUE!`, as it is in LibreOffice.
fn least_squares(d: &Design, y: &[f64], constant: bool) -> Res<Fit> {
    let n = y.len();
    let k = d.x.len();
    let p = k + usize::from(constant);
    if n < p || n < 2 {
        return Err(DOMAIN);
    }
    if k == 1 {
        return single_regressor(&d.x[0], y, constant);
    }
    // Columns of the design matrix, then the reflections.
    let mut cols: Vec<Vec<f64>> = d.x.clone();
    if constant {
        cols.push(vec![1.0; n]);
    }
    let norms: Vec<f64> = cols
        .iter()
        .map(|c| c.iter().map(|v| v * v).sum::<f64>().sqrt())
        .collect();
    let mut qty = y.to_vec();
    let mut r = vec![vec![0.0; p]; p];
    for j in 0..p {
        let norm: f64 = cols[j][j..].iter().map(|v| v * v).sum::<f64>().sqrt();
        if norm <= 1e-10 * norms[j] || norms[j] == 0.0 {
            return Err(ErrorKind::Value);
        }
        let alpha = if cols[j][j] > 0.0 { -norm } else { norm };
        let mut v: Vec<f64> = cols[j][j..].to_vec();
        v[0] -= alpha;
        let vnorm2: f64 = v.iter().map(|t| t * t).sum();
        r[j][j] = alpha;
        if vnorm2 > 0.0 {
            for col in cols.iter_mut().skip(j + 1) {
                let dot: f64 = (j..n).map(|i| v[i - j] * col[i]).sum();
                let f = 2.0 * dot / vnorm2;
                for i in j..n {
                    col[i] -= f * v[i - j];
                }
            }
            let dot: f64 = (j..n).map(|i| v[i - j] * qty[i]).sum();
            let f = 2.0 * dot / vnorm2;
            for i in j..n {
                qty[i] -= f * v[i - j];
            }
        }
        for c in j + 1..p {
            r[j][c] = cols[c][j];
        }
    }
    // Back substitution.
    let mut beta = vec![0.0; p];
    for j in (0..p).rev() {
        let mut s = qty[j];
        for c in j + 1..p {
            s -= r[j][c] * beta[c];
        }
        beta[j] = s / r[j][j];
    }
    // The inverse of R, whose rows give `(X'X)^-1 = R^-1 R^-T`.
    let mut rinv = vec![vec![0.0; p]; p];
    for c in 0..p {
        rinv[c][c] = 1.0 / r[c][c];
        for j in (0..c).rev() {
            let mut s = 0.0;
            for m in j + 1..=c {
                s += r[j][m] * rinv[m][c];
            }
            rinv[j][c] = -s / r[j][j];
        }
    }
    let inv_diag: Vec<f64> = (0..p)
        .map(|j| (j..p).map(|c| rinv[j][c] * rinv[j][c]).sum())
        .collect();
    Ok(finish(d.x.as_slice(), y, beta, inv_diag, constant))
}

/// The fit of one regressor, by the means as LibreOffice does it (so that data on a line, with
/// numbers that are exact, leave no residual at all).
fn single_regressor(x: &[f64], y: &[f64], constant: bool) -> Res<Fit> {
    let n = y.len() as f64;
    if constant {
        let (mx, my) = (x.iter().sum::<f64>() / n, y.iter().sum::<f64>() / n);
        let sxx: f64 = x.iter().map(|v| (v - mx).powi(2)).sum();
        let sxy: f64 = x.iter().zip(y).map(|(a, b)| (a - mx) * (b - my)).sum();
        let sq: f64 = x.iter().map(|v| v * v).sum();
        if sxx <= 1e-20 * sq {
            return Err(ErrorKind::Value);
        }
        let m = sxy / sxx;
        let b = my - m * mx;
        Ok(finish(
            &[x.to_vec()],
            y,
            vec![m, b],
            vec![1.0 / sxx, 1.0 / n + mx * mx / sxx],
            true,
        ))
    } else {
        let sxx: f64 = x.iter().map(|v| v * v).sum();
        if sxx == 0.0 {
            return Err(ErrorKind::Value);
        }
        let m = x.iter().zip(y).map(|(a, b)| a * b).sum::<f64>() / sxx;
        Ok(finish(&[x.to_vec()], y, vec![m], vec![1.0 / sxx], false))
    }
}

/// The sums of squares of a fit.
fn finish(x: &[Vec<f64>], y: &[f64], beta: Vec<f64>, inv_diag: Vec<f64>, constant: bool) -> Fit {
    let (n, k) = (y.len(), x.len());
    let p = k + usize::from(constant);
    let mean = y.iter().sum::<f64>() / n as f64;
    let (mut ssresid, mut ssreg) = (0.0, 0.0);
    for i in 0..n {
        let mut fitted = if constant { beta[k] } else { 0.0 };
        for (j, col) in x.iter().enumerate() {
            fitted += beta[j] * col[i];
        }
        ssresid += (y[i] - fitted).powi(2);
        ssreg += if constant {
            (fitted - mean).powi(2)
        } else {
            fitted * fitted
        };
    }
    // A fit with nothing left over to estimate has no residual; one that is perfect up to the
    // rounding of the computation has a noise that LibreOffice's own rounding does not repeat.
    let scale: f64 = y.iter().map(|v| v * v).sum();
    let mut noise = false;
    if n == p {
        ssresid = 0.0;
    } else if ssresid != 0.0 && ssresid <= 1e-26 * scale {
        noise = true;
    }
    Fit {
        noise,
        beta,
        inv_diag,
        ssresid,
        ssreg,
        df: (n - p) as f64,
        constant,
    }
}

fn number_cell(v: f64) -> Value {
    if v.is_finite() {
        Value::Number(v)
    } else {
        Value::Error(ErrorKind::Num)
    }
}

/// The `LINEST` array of a fit: the coefficients (the last regressor first, the intercept last),
/// and with `stats` the errors, `r²` and the standard error of `y`, `F` and the degrees of
/// freedom, and the sums of squares.
fn estimates(fit: &Fit, k: usize, stats: bool, exp: bool) -> Value {
    let cols = k + 1;
    let coef = |j: usize| -> f64 {
        let b = if j < k {
            fit.beta[j]
        } else if fit.constant {
            fit.beta[k]
        } else {
            0.0
        };
        if exp { b.exp() } else { b }
    };
    // Position 0 is the last regressor; the intercept is at `k`.
    let order = |pos: usize| if pos < k { k - 1 - pos } else { k };
    let mut data: Vec<Value> = (0..cols).map(|p| number_cell(coef(order(p)))).collect();
    if !stats {
        return array(1, cols, data);
    }
    let sey = if fit.df > 0.0 {
        (fit.ssresid / fit.df).sqrt()
    } else {
        0.0
    };
    let se = |j: usize| -> Value {
        if j == k && !fit.constant {
            return na();
        }
        number_cell(fit.inv_diag[j].sqrt() * sey)
    };
    data.extend((0..cols).map(|p| se(order(p))));
    let sstot = fit.ssreg + fit.ssresid;
    let r2 = if fit.ssresid == 0.0 || sstot == 0.0 {
        1.0
    } else {
        fit.ssreg / sstot
    };
    let f = if fit.df > 0.0 && fit.ssresid > 0.0 {
        number_cell((fit.ssreg / k as f64) / (fit.ssresid / fit.df))
    } else {
        na()
    };
    let mut row = |a: Value, b: Value| {
        data.push(a);
        data.push(b);
        data.extend((2..cols).map(|_| na()));
    };
    row(number_cell(r2), number_cell(sey));
    row(f, number_cell(fit.df));
    row(number_cell(fit.ssreg), number_cell(fit.ssresid));
    array(5, cols, data)
}

fn linest(e: &mut Engine, a: &[Value], exp: bool) -> Res<Value> {
    let x = opt(a, 1);
    let constant = flag(e, a, 2, true)?;
    let stats = flag(e, a, 3, false)?;
    let d = design(e, &a[0], x)?;
    let y: Vec<f64> = if exp {
        if d.y.iter().any(|v| *v <= 0.0) {
            return Err(DOMAIN);
        }
        d.y.iter().map(|v| v.ln()).collect()
    } else {
        d.y.clone()
    };
    let fit = least_squares(&d, &y, constant)?;
    if stats && fit.noise {
        crate::notes::note(crate::notes::PERFECT_FIT);
        return Err(ErrorKind::Unsupported);
    }
    Ok(estimates(&fit, d.x.len(), stats, exp))
}

fn trend(e: &mut Engine, a: &[Value], exp: bool) -> Res<Value> {
    // Text that is not a boolean gives an array of `#VALUE!` of the shape of the answer.
    let mut bad_const = false;
    let constant = match opt(a, 3) {
        None => true,
        Some(v) => match e.scalar(v.clone()) {
            Value::Error(k) => return Err(k),
            v => crate::coerce::to_bool(&v, &e.opts.locale).unwrap_or_else(|_| {
                bad_const = true;
                true
            }),
        },
    };
    let d = design(e, &a[0], opt(a, 1))?;
    let y: Vec<f64> = if exp {
        if d.y.iter().any(|v| *v <= 0.0) {
            return Err(DOMAIN);
        }
        d.y.iter().map(|v| v.ln()).collect()
    } else {
        d.y.clone()
    };
    let fit = least_squares(&d, &y, constant)?;
    let k = d.x.len();
    let predict = |point: &[f64]| -> Value {
        let mut v = if constant { fit.beta[k] } else { 0.0 };
        for (j, x) in point.iter().enumerate() {
            v += fit.beta[j] * x;
        }
        if bad_const {
            return Value::Error(ErrorKind::Value);
        }
        number_cell(if exp { v.exp() } else { v })
    };
    match opt(a, 2) {
        None => {
            let n = d.y.len();
            let out: Vec<Value> = (0..n)
                .map(|i| predict(&d.x.iter().map(|c| c[i]).collect::<Vec<_>>()))
                .collect();
            Ok(if d.by_rows {
                array(1, n, out)
            } else {
                array(n, 1, out)
            })
        }
        Some(nx) => {
            let m = real_matrix(e, nx)?;
            if d.single {
                let out: Vec<Value> = m.data.iter().map(|v| predict(&[*v])).collect();
                return Ok(array(m.rows, m.cols, out));
            }
            if d.by_rows {
                if m.rows != k {
                    return Err(DOMAIN);
                }
                let out: Vec<Value> = (0..m.cols)
                    .map(|c| predict(&(0..k).map(|r| m.at(r, c)).collect::<Vec<_>>()))
                    .collect();
                Ok(array(1, m.cols, out))
            } else {
                if m.cols != k {
                    return Err(DOMAIN);
                }
                let out: Vec<Value> = (0..m.rows)
                    .map(|r| predict(&(0..k).map(|c| m.at(r, c)).collect::<Vec<_>>()))
                    .collect();
                Ok(array(m.rows, 1, out))
            }
        }
    }
}

/// The numbers of an argument of `FREQUENCY`: the numbers and booleans, an error being that error
/// and text or blanks being skipped.
fn frequency_values(e: &mut Engine, v: &Value) -> Res<Vec<f64>> {
    let m = e.array_of(v.clone());
    let mut out = Vec::new();
    for x in &m.data {
        match x {
            Value::Number(n) => out.push(*n),
            Value::Bool(b) => out.push(f64::from(u8::from(*b))),
            Value::Error(k) => return Err(*k),
            _ => {}
        }
    }
    Ok(out)
}

fn frequency(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let data = frequency_values(e, &a[0])?;
    let bins = frequency_values(e, &a[1])?;
    if data.is_empty() || bins.is_empty() {
        return Err(ErrorKind::Value);
    }
    let mut order: Vec<usize> = (0..bins.len()).collect();
    order.sort_by(|&p, &q| bins[p].total_cmp(&bins[q]));
    // Which of equal bins gets the count depends on how LibreOffice sorts them.
    if order.windows(2).any(|w| bins[w[0]] == bins[w[1]]) {
        crate::notes::note(crate::notes::FREQUENCY_EQUAL_BINS);
        return Err(ErrorKind::Unsupported);
    }
    let sorted: Vec<f64> = order.iter().map(|&i| bins[i]).collect();
    let mut counts = vec![0.0; bins.len() + 1];
    for x in &data {
        let p = sorted.partition_point(|b| b < x);
        if p == sorted.len() {
            counts[bins.len()] += 1.0;
            continue;
        }
        counts[order[p]] += 1.0;
    }
    Ok(array(
        counts.len(),
        1,
        counts.into_iter().map(Value::Number).collect(),
    ))
}

fn mode_mult(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let numbers = e.numbers(a)?;
    let mut seen: Vec<(f64, usize)> = Vec::new();
    for x in &numbers {
        match seen.iter_mut().find(|(v, _)| v == x) {
            Some((_, c)) => *c += 1,
            None => seen.push((*x, 1)),
        }
    }
    let top = seen.iter().map(|(_, c)| *c).max().unwrap_or(0);
    if top < 2 {
        return Err(ErrorKind::Value);
    }
    let modes: Vec<Value> = seen
        .iter()
        .filter(|(_, c)| *c == top)
        .map(|(v, _)| Value::Number(*v))
        .collect();
    Ok(array(modes.len(), 1, modes))
}

pub(super) fn register(r: &mut Reg) {
    r.eager("LINEST", 1, 4, |e, a| linest(e, a, false))
        .array_args();
    r.eager("LOGEST", 1, 4, |e, a| linest(e, a, true))
        .array_args();
    r.eager("TREND", 1, 4, |e, a| trend(e, a, false))
        .array_args();
    r.eager("GROWTH", 1, 4, |e, a| trend(e, a, true))
        .array_args();
    r.eager("FREQUENCY", 2, 2, frequency).array_args();
    r.eager("MODE.MULT", 1, VAR, mode_mult).array_args();
}
