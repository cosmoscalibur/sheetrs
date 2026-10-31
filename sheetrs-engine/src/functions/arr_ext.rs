//! Matrices, `LOOKUP`, sums of squares and `TYPE`.

use super::lookup::position;
use super::stats::values_equal;
use super::{Reg, VAR, num, opt};
use crate::engine::{Engine, Res};
use crate::value::{Array, ErrorKind, Value};
use std::sync::Arc;

/// A numeric matrix; anything that is not a number is `#VALUE!`.
fn matrix(e: &mut Engine, v: &Value) -> Res<(usize, usize, Vec<f64>)> {
    let a = e.array_of(v.clone());
    let mut out = Vec::with_capacity(a.data.len());
    for x in &a.data {
        match x {
            Value::Number(n) => out.push(*n),
            Value::Error(k) => return Err(*k),
            _ => return Err(ErrorKind::Value),
        }
    }
    Ok((a.rows as usize, a.cols as usize, out))
}

fn array(rows: usize, cols: usize, data: Vec<f64>) -> Value {
    Value::Array(Arc::new(Array::new(
        rows as u32,
        cols as u32,
        data.into_iter().map(Value::Number).collect(),
    )))
}

fn pair_sums(e: &mut Engine, a: &[Value], f: fn(f64, f64) -> f64) -> Res<Value> {
    let (x, y) = (e.array_of(a[0].clone()), e.array_of(a[1].clone()));
    if x.data.len() != y.data.len() {
        return Err(ErrorKind::Na);
    }
    let mut s = 0.0;
    for (p, q) in x.data.iter().zip(&y.data) {
        if let Value::Error(k) = p {
            return Err(*k);
        }
        if let Value::Error(k) = q {
            return Err(*k);
        }
        if let (Value::Number(p), Value::Number(q)) = (p, q) {
            s += f(*p, *q);
        }
    }
    num(s)
}

fn determinant(n: usize, mut m: Vec<f64>) -> f64 {
    let mut det = 1.0;
    for c in 0..n {
        let p = (c..n)
            .max_by(|&i, &j| m[i * n + c].abs().total_cmp(&m[j * n + c].abs()))
            .unwrap_or(c);
        if m[p * n + c] == 0.0 {
            return 0.0;
        }
        if p != c {
            for k in 0..n {
                m.swap(p * n + k, c * n + k);
            }
            det = -det;
        }
        det *= m[c * n + c];
        for r in c + 1..n {
            let f = m[r * n + c] / m[c * n + c];
            for k in c..n {
                m[r * n + k] -= f * m[c * n + k];
            }
        }
    }
    det
}

fn inverse(n: usize, m: &[f64]) -> Option<Vec<f64>> {
    let mut a = m.to_vec();
    let mut inv: Vec<f64> = (0..n * n)
        .map(|i| f64::from(u8::from(i / n == i % n)))
        .collect();
    for c in 0..n {
        let p = (c..n).max_by(|&i, &j| a[i * n + c].abs().total_cmp(&a[j * n + c].abs()))?;
        if a[p * n + c].abs() < 1e-300 {
            return None;
        }
        for k in 0..n {
            a.swap(p * n + k, c * n + k);
            inv.swap(p * n + k, c * n + k);
        }
        let d = a[c * n + c];
        for k in 0..n {
            a[c * n + k] /= d;
            inv[c * n + k] /= d;
        }
        for r in 0..n {
            if r != c {
                let f = a[r * n + c];
                for k in 0..n {
                    a[r * n + k] -= f * a[c * n + k];
                    inv[r * n + k] -= f * inv[c * n + k];
                }
            }
        }
    }
    Some(inv)
}

pub(super) fn register(r: &mut Reg) {
    // UNIQUE(array, [by_col], [exactly_once]): the distinct rows (or columns), in
    // the order they first appear; text compares without regard to case.
    r.eager("UNIQUE", 1, 3, |e, a| {
        let m = e.array_of(a[0].clone());
        let by_col = match opt(a, 1) {
            Some(v) => e.boolean(v)?,
            None => false,
        };
        let once = match opt(a, 2) {
            Some(v) => e.boolean(v)?,
            None => false,
        };
        let (lanes, depth) = if by_col {
            (m.cols, m.rows)
        } else {
            (m.rows, m.cols)
        };
        let cell = |lane: u32, i: u32| {
            if by_col {
                m.get(i, lane)
            } else {
                m.get(lane, i)
            }
        };
        let same = |e: &mut Engine, x: u32, y: u32| {
            (0..depth).all(|i| {
                let (p, q) = (cell(x, i), cell(y, i));
                match (p, q) {
                    (Value::Empty, Value::Empty) => true,
                    _ => values_equal(p, q, e.collator(false)),
                }
            })
        };
        let mut kept: Vec<(u32, u32)> = Vec::new(); // (first lane, occurrences)
        for lane in 0..lanes {
            match (0..kept.len()).find(|&k| same(e, kept[k].0, lane)) {
                Some(k) => kept[k].1 += 1,
                None => kept.push((lane, 1)),
            }
        }
        let picked: Vec<u32> = kept
            .iter()
            .filter(|(_, n)| !once || *n == 1)
            .map(|(l, _)| *l)
            .collect();
        if picked.is_empty() {
            return Err(ErrorKind::Value);
        }
        let n = picked.len() as u32;
        let mut out = Vec::with_capacity((n * depth) as usize);
        if by_col {
            for r in 0..depth {
                for &l in &picked {
                    out.push(cell(l, r).clone());
                }
            }
            Ok(Value::Array(Arc::new(Array::new(depth, n, out))))
        } else {
            for &l in &picked {
                for i in 0..depth {
                    out.push(cell(l, i).clone());
                }
            }
            Ok(Value::Array(Arc::new(Array::new(n, depth, out))))
        }
    })
    .array_args();
    r.eager("TRANSPOSE", 1, 1, |e, a| {
        let m = e.array_of(a[0].clone());
        let (rows, cols) = (m.rows, m.cols);
        let mut out = Vec::with_capacity(m.data.len());
        for c in 0..cols {
            for r in 0..rows {
                out.push(m.get(r, c).clone());
            }
        }
        Ok(Value::Array(Arc::new(Array::new(cols, rows, out))))
    })
    .array_args();
    r.eager("MMULT", 2, 2, |e, a| {
        let ((ar, ac, x), (br, bc, y)) = (matrix(e, &a[0])?, matrix(e, &a[1])?);
        if ac != br {
            return Err(ErrorKind::Value);
        }
        let mut out = vec![0.0; ar * bc];
        for i in 0..ar {
            for j in 0..bc {
                out[i * bc + j] = (0..ac).map(|k| x[i * ac + k] * y[k * bc + j]).sum();
            }
        }
        Ok(array(ar, bc, out))
    })
    .array_args();
    r.eager("MDETERM", 1, 1, |e, a| {
        let (rows, cols, m) = matrix(e, &a[0])?;
        if rows != cols || rows == 0 {
            return Err(ErrorKind::Value);
        }
        num(determinant(rows, m))
    })
    .array_args();
    r.eager("MINVERSE", 1, 1, |e, a| {
        let (rows, cols, m) = matrix(e, &a[0])?;
        if rows != cols || rows == 0 {
            return Err(ErrorKind::Value);
        }
        inverse(rows, &m)
            .map(|v| array(rows, rows, v))
            .ok_or(ErrorKind::Domain)
    })
    .array_args();
    r.eager("MUNIT", 1, 1, |e, a| {
        let n = e.num(&a[0])?.trunc();
        if n < 1.0 {
            return Err(ErrorKind::Value);
        }
        let n = n as usize;
        Ok(array(
            n,
            n,
            (0..n * n)
                .map(|i| f64::from(u8::from(i / n == i % n)))
                .collect(),
        ))
    });
    r.eager("SUMXMY2", 2, 2, |e, a| {
        pair_sums(e, a, |x, y| (x - y) * (x - y))
    })
    .array_args();
    r.eager("SUMX2MY2", 2, 2, |e, a| {
        pair_sums(e, a, |x, y| x * x - y * y)
    })
    .array_args();
    r.eager("SUMX2PY2", 2, 2, |e, a| {
        pair_sums(e, a, |x, y| x * x + y * y)
    })
    .array_args();

    r.eager("LOOKUP", 2, 3, |e, a| {
        let needle = e.scalar(a[0].clone());
        let probe = e.array_of(a[1].clone());
        // An array that is not a range is searched as LibreOffice does.
        if let Value::Array(m) = &a[1] {
            if let Value::Error(k) = needle {
                return Err(k);
            }
            let (lane, vertical) = super::matrix_lookup::lane(m);
            let numeric = m
                .data
                .iter()
                .all(|v| matches!(v, Value::Number(_) | Value::Bool(_) | Value::Error(_)));
            let i =
                super::matrix_lookup::lookup_position(&needle, &lane, numeric, e.collator(false))
                    .ok_or(ErrorKind::Na)?;
            return match opt(a, 2) {
                Some(res) => e
                    .array_of(res.clone())
                    .data
                    .get(i)
                    .cloned()
                    .ok_or(ErrorKind::Na),
                None if vertical => Ok(m.get(i as u32, m.cols - 1).clone()),
                None => Ok(m.get(m.rows - 1, i as u32).clone()),
            };
        }
        if let Some(res) = opt(a, 2) {
            let result = e.array_of(res.clone());
            let i = position(&needle, &probe.data, 1.0, &e.opts, e.collator(false))
                .ok_or(ErrorKind::Na)?;
            return result.data.get(i).cloned().ok_or(ErrorKind::Na);
        }
        // Array form: the longer side is searched, the last row or column answers.
        let (rows, cols) = (probe.rows as usize, probe.cols as usize);
        if cols > rows {
            let first: Vec<Value> = (0..cols).map(|c| probe.get(0, c as u32).clone()).collect();
            let i =
                position(&needle, &first, 1.0, &e.opts, e.collator(false)).ok_or(ErrorKind::Na)?;
            Ok(probe.get(rows as u32 - 1, i as u32).clone())
        } else {
            let first: Vec<Value> = (0..rows).map(|r| probe.get(r as u32, 0).clone()).collect();
            let i =
                position(&needle, &first, 1.0, &e.opts, e.collator(false)).ok_or(ErrorKind::Na)?;
            Ok(probe.get(i as u32, cols as u32 - 1).clone())
        }
    })
    .array_args();
    r.eager("TYPE", 1, 1, |e, a| {
        let v = match &a[0] {
            // It does not look inside a range.
            Value::Ref(r) if !r.is_single() => return Err(ErrorKind::Value),
            v => e.scalar(v.clone()),
        };
        num(match v {
            Value::Number(_) | Value::Empty | Value::Ref(_) => 1.0,
            Value::Text(_) => 2.0,
            Value::Bool(_) => 1.0,
            Value::Error(_) => 16.0,
            Value::Array(_) => 64.0,
        })
    });
    let _ = VAR;
}
