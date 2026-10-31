//! The array functions of the dynamic-array era and the text functions that split and cut:
//! `SEQUENCE`, `FILTER`, `SORT`, `SORTBY`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `HSTACK`,
//! `VSTACK`, `EXPAND`, `TOCOL`, `TOROW`, `WRAPCOLS`, `WRAPROWS`, `RANDARRAY`, `TEXTSPLIT`,
//! `TEXTBEFORE`, `TEXTAFTER`, `ENCODEURL` and `ROT13`. What they answer for an argument outside
//! their domain is what the oracle (`tests/assets/semantics.*`) has from LibreOffice 26.8.

use super::{Reg, VAR, opt, opt_num};
use crate::engine::{Engine, Res};
use crate::value::{Array, ErrorKind, Value, compare};
use std::cmp::Ordering;
use std::fmt::Write;
use std::sync::Arc;

const DOMAIN: ErrorKind = ErrorKind::Domain;

fn array(rows: usize, cols: usize, data: Vec<Value>) -> Value {
    Value::Array(Arc::new(Array::new(rows as u32, cols as u32, data)))
}

fn na() -> Value {
    Value::Error(ErrorKind::Na)
}

/// An optional whole-number argument: `None` when it is absent or omitted.
fn whole(e: &mut Engine, a: &[Value], i: usize) -> Res<Option<i64>> {
    match opt(a, i) {
        Some(v) => Ok(Some(e.num(v)?.trunc() as i64)),
        None => Ok(None),
    }
}

fn flag(e: &mut Engine, a: &[Value], i: usize) -> Res<bool> {
    match opt(a, i) {
        Some(v) => e.boolean(v),
        None => Ok(false),
    }
}

/// The rows of `m` at `which`, in that order.
fn pick_rows(m: &Array, which: &[usize]) -> Value {
    let mut out = Vec::with_capacity(which.len() * m.cols as usize);
    for &r in which {
        for c in 0..m.cols {
            out.push(m.get(r as u32, c).clone());
        }
    }
    array(which.len(), m.cols as usize, out)
}

fn pick_cols(m: &Array, which: &[usize]) -> Value {
    let mut out = Vec::with_capacity(which.len() * m.rows as usize);
    for r in 0..m.rows {
        for &c in which {
            out.push(m.get(r, c as u32).clone());
        }
    }
    array(m.rows as usize, which.len(), out)
}

/// `SORT` and `SORTBY` of a range give 0 where the range has an error (an array built by another
/// function keeps its errors).
fn errors_of_a_range_are_zero(source: &Value, sorted: Value) -> Value {
    match (source, sorted) {
        (Value::Ref(_), Value::Array(m)) => {
            let data = m
                .data
                .iter()
                .map(|v| match v {
                    Value::Error(_) => Value::Number(0.0),
                    v => v.clone(),
                })
                .collect();
            Value::Array(Arc::new(Array::new(m.rows, m.cols, data)))
        }
        (_, sorted) => sorted,
    }
}

/// Keys with an error give an order that LibreOffice does not derive from the keys.
fn unordered_errors(key: &[Value]) -> Res<()> {
    if key.iter().any(|v| matches!(v, Value::Error(_))) {
        crate::notes::note(crate::notes::SORT_ERROR_KEY);
        return Err(ErrorKind::Unsupported);
    }
    Ok(())
}

/// Whether a `FILTER` condition holds: numbers and booleans do, text is an error.
fn truth(v: &Value) -> Res<bool> {
    match v {
        Value::Number(n) => Ok(*n != 0.0),
        Value::Bool(b) => Ok(*b),
        Value::Error(k) => Err(*k),
        // Text, and a blank cell: LibreOffice does not take them for a condition.
        _ => Err(DOMAIN),
    }
}

/// The lanes (rows, or columns when `by_col`) of `m` in the order of the keys, ascending or
/// descending; the sort is stable, text compares without regard to case.
fn sorted_lanes(e: &Engine, keys: &[Vec<Value>], descending: &[bool], n: usize) -> Vec<usize> {
    let col = e.collator(false);
    let mut order: Vec<usize> = (0..n).collect();
    order.sort_by(|&x, &y| {
        for (k, desc) in keys.iter().zip(descending) {
            // Blank cells go last, whichever the order; an error is the greatest value.
            let o = match (&k[x], &k[y]) {
                (Value::Empty, Value::Empty) => Ordering::Equal,
                (Value::Empty, _) => return Ordering::Greater,
                (_, Value::Empty) => return Ordering::Less,
                (p, q) => {
                    let o = match (p, q) {
                        (Value::Error(_), Value::Error(_)) => Ordering::Equal,
                        (Value::Error(_), _) => Ordering::Greater,
                        (_, Value::Error(_)) => Ordering::Less,
                        _ => compare(p, q, col),
                    };
                    if *desc { o.reverse() } else { o }
                }
            };
            if o != Ordering::Equal {
                return o;
            }
        }
        Ordering::Equal
    });
    order
}

fn order_arg(e: &mut Engine, a: &[Value], i: usize) -> Res<bool> {
    match whole(e, a, i)? {
        None | Some(1) => Ok(false),
        Some(-1) => Ok(true),
        _ => Err(DOMAIN),
    }
}

// ----- the delimiters of the text functions -----

fn delimiters(e: &mut Engine, v: &Value) -> Res<Vec<String>> {
    let m = e.array_of(v.clone());
    let mut out = Vec::new();
    for x in &m.data {
        if let Value::Error(k) = x {
            return Err(*k);
        }
        let t = e.text(x)?;
        if !t.is_empty() {
            out.push(t);
        }
    }
    Ok(out)
}

fn fold_char(c: char) -> String {
    let mut b = [0u8; 4];
    crate::casemap::fold_search(c.encode_utf8(&mut b)).0
}

/// The end of `delim` in `text` if it starts at `pos`.
fn match_at(text: &str, pos: usize, delim: &str, folded: &str, ci: bool) -> Option<usize> {
    if !ci {
        return text[pos..].starts_with(delim).then(|| pos + delim.len());
    }
    let mut acc = String::new();
    let mut end = pos;
    for c in text[pos..].chars() {
        acc.push_str(&fold_char(c));
        end += c.len_utf8();
        if acc.len() >= folded.len() {
            return (acc == folded).then_some(end);
        }
    }
    None
}

/// Every place a delimiter starts, `(start, end)`, longest first at the same start.
fn candidates(text: &str, delims: &[String], ci: bool) -> Vec<(usize, usize)> {
    let folded: Vec<String> = delims
        .iter()
        .map(|d| {
            if ci {
                crate::casemap::fold_search(d).0
            } else {
                String::new()
            }
        })
        .collect();
    let mut out = Vec::new();
    for (pos, _) in text.char_indices() {
        let mut best: Option<usize> = None;
        for (d, f) in delims.iter().zip(&folded) {
            if let Some(end) = match_at(text, pos, d, f, ci) {
                best = best.max(Some(end));
            }
        }
        if let Some(end) = best {
            out.push((pos, end));
        }
    }
    out
}

/// The delimiters found from the left without overlapping.
fn forward(text: &str, delims: &[String], ci: bool) -> Vec<(usize, usize)> {
    let mut out = Vec::new();
    let mut from = 0;
    for (s, e) in candidates(text, delims, ci) {
        if s >= from {
            out.push((s, e));
            from = e;
        }
    }
    out
}

/// The delimiters found from the right without overlapping, the last one first.
fn backward(text: &str, delims: &[String], ci: bool) -> Vec<(usize, usize)> {
    let mut all = candidates(text, delims, ci);
    all.sort_by(|x, y| y.1.cmp(&x.1).then(x.0.cmp(&y.0)));
    let mut out = Vec::new();
    let mut limit = text.len();
    for (s, e) in all {
        if e <= limit {
            out.push((s, e));
            limit = s;
        }
    }
    out
}

fn pieces(text: &str, delims: &[String], ci: bool) -> Vec<String> {
    let mut out = Vec::new();
    let mut from = 0;
    for (s, e) in forward(text, delims, ci) {
        out.push(text[from..s].to_string());
        from = e;
    }
    out.push(text[from..].to_string());
    out
}

/// The match mode: zero is case-sensitive, any other number is not.
fn match_mode(e: &mut Engine, a: &[Value], i: usize) -> Res<bool> {
    Ok(whole(e, a, i)?.is_some_and(|m| m != 0))
}

/// `TEXTBEFORE` (`before`) and `TEXTAFTER`.
fn text_around(e: &mut Engine, a: &[Value], before: bool) -> Res<Value> {
    let text = e.text(&a[0])?;
    let delims = delimiters(e, &a[1])?;
    let instance = whole(e, a, 2)?.unwrap_or(1);
    let ci = match_mode(e, a, 3)?;
    let match_end = whole(e, a, 4)?.is_some_and(|m| m != 0);
    let not_found = || opt(a, 5).cloned().unwrap_or_else(na);
    if text.is_empty() {
        return Err(DOMAIN);
    }
    if delims.is_empty() || instance == 0 {
        return Ok(na());
    }
    let found = if instance > 0 {
        forward(&text, &delims, ci)
    } else {
        backward(&text, &delims, ci)
    };
    let k = instance.unsigned_abs() as usize;
    let at = if k <= found.len() {
        Some(found[k - 1])
    } else if match_end && k == found.len() + 1 {
        // The end of the text (the start, counting from the right) is a delimiter too.
        Some(if instance > 0 {
            (text.len(), text.len())
        } else {
            (0, 0)
        })
    } else {
        None
    };
    Ok(match at {
        None => not_found(),
        Some((s, en)) => Value::text(if before { &text[..s] } else { &text[en..] }),
    })
}

fn url_encode(s: &str) -> String {
    let mut out = String::new();
    for b in s.bytes() {
        if b.is_ascii_alphanumeric() || b == b'-' || b == b'_' {
            out.push(char::from(b));
        } else {
            let _ = write!(out, "%{b:02X}");
        }
    }
    out
}

pub(super) fn register(r: &mut Reg) {
    r.eager("SEQUENCE", 1, 4, |e, a| {
        let rows = e.num(&a[0])?.trunc();
        let cols = opt_num(e, a, 1, 1.0)?.trunc();
        let (start, step) = (opt_num(e, a, 2, 1.0)?, opt_num(e, a, 3, 1.0)?);
        if rows < 1.0 || cols < 1.0 {
            return Err(DOMAIN);
        }
        let (rows, cols) = (rows as usize, cols as usize);
        let data = (0..rows * cols)
            .map(|i| Value::Number(start + step * i as f64))
            .collect();
        Ok(array(rows, cols, data))
    });

    r.eager("FILTER", 2, 3, |e, a| {
        let m = e.array_of(a[0].clone());
        let inc = e.array_of(a[1].clone());
        let by_rows = inc.cols == 1 && inc.rows == m.rows;
        let by_cols = !by_rows && inc.rows == 1 && inc.cols == m.cols;
        if !by_rows && !by_cols {
            return Err(DOMAIN);
        }
        let mut keep = Vec::new();
        for (i, v) in inc.data.iter().enumerate() {
            if truth(v)? {
                keep.push(i);
            }
        }
        if keep.is_empty() {
            return match opt(a, 2) {
                Some(v) => Ok(v.clone()),
                None => Err(DOMAIN),
            };
        }
        Ok(if by_rows {
            pick_rows(&m, &keep)
        } else {
            pick_cols(&m, &keep)
        })
    })
    .array_args();

    r.eager("SORT", 1, 4, |e, a| {
        let m = e.array_of(a[0].clone());
        let index = whole(e, a, 1)?.unwrap_or(1);
        let desc = order_arg(e, a, 2)?;
        let by_col = flag(e, a, 3)?;
        let (lanes, depth) = if by_col {
            (m.cols, m.rows)
        } else {
            (m.rows, m.cols)
        };
        if index < 1 || index > i64::from(depth) {
            return Err(DOMAIN);
        }
        let key: Vec<Value> = (0..lanes)
            .map(|l| {
                if by_col {
                    m.get(index as u32 - 1, l).clone()
                } else {
                    m.get(l, index as u32 - 1).clone()
                }
            })
            .collect();
        if !matches!(a[0], Value::Ref(_)) {
            unordered_errors(&key)?;
        }
        let order = sorted_lanes(e, &[key], &[desc], lanes as usize);
        let sorted = if by_col {
            pick_cols(&m, &order)
        } else {
            pick_rows(&m, &order)
        };
        Ok(errors_of_a_range_are_zero(&a[0], sorted))
    })
    .array_args();

    r.eager("SORTBY", 2, VAR, |e, a| {
        let m = e.array_of(a[0].clone());
        let mut keys = Vec::new();
        let mut desc = Vec::new();
        let mut by_col: Option<bool> = None;
        let mut i = 1;
        while i < a.len() {
            let k = e.array_of(a[i].clone());
            let cols = if k.cols == 1 && k.rows == m.rows {
                false
            } else if k.rows == 1 && k.cols == m.cols {
                true
            } else {
                return Err(DOMAIN);
            };
            if by_col.is_some_and(|b| b != cols) {
                return Err(DOMAIN);
            }
            by_col = Some(cols);
            unordered_errors(&k.data)?;
            keys.push(k.data.clone());
            desc.push(order_arg(e, a, i + 1)?);
            i += 2;
        }
        let lanes = if by_col == Some(true) { m.cols } else { m.rows } as usize;
        let order = sorted_lanes(e, &keys, &desc, lanes);
        let sorted = if by_col == Some(true) {
            pick_cols(&m, &order)
        } else {
            pick_rows(&m, &order)
        };
        Ok(errors_of_a_range_are_zero(&a[0], sorted))
    })
    .array_args();

    r.eager("TAKE", 2, 3, |e, a| {
        let m = e.array_of(a[0].clone());
        let span = |n: Option<i64>, len: u32| -> (usize, usize) {
            match n {
                None => (0, len as usize),
                Some(n) if n >= 0 => (0, (n as usize).min(len as usize)),
                Some(n) => {
                    let k = (n.unsigned_abs() as usize).min(len as usize);
                    (len as usize - k, len as usize)
                }
            }
        };
        let (r0, r1) = span(whole(e, a, 1)?, m.rows);
        let (c0, c1) = span(whole(e, a, 2)?, m.cols);
        if r0 == r1 || c0 == c1 {
            return Err(ErrorKind::Na);
        }
        Ok(window(&m, r0, r1, c0, c1))
    })
    .array_args();

    r.eager("DROP", 2, 3, |e, a| {
        let m = e.array_of(a[0].clone());
        let span = |n: Option<i64>, len: u32| -> (usize, usize) {
            let len = len as usize;
            match n {
                None => (0, len),
                Some(n) if n >= 0 => ((n as usize).min(len), len),
                Some(n) => (0, len - (n.unsigned_abs() as usize).min(len)),
            }
        };
        let (r0, r1) = span(whole(e, a, 1)?, m.rows);
        let (c0, c1) = span(whole(e, a, 2)?, m.cols);
        if r0 >= r1 || c0 >= c1 {
            return Err(DOMAIN);
        }
        Ok(window(&m, r0, r1, c0, c1))
    })
    .array_args();

    for (name, cols) in [("CHOOSECOLS", true), ("CHOOSEROWS", false)] {
        let f: super::EagerFn = if cols {
            |e, a| choose(e, a, true)
        } else {
            |e, a| choose(e, a, false)
        };
        r.eager(name, 2, VAR, f).array_args();
    }

    r.eager("HSTACK", 1, VAR, |e, a| stack(e, a, true))
        .array_args();
    r.eager("VSTACK", 1, VAR, |e, a| stack(e, a, false))
        .array_args();

    r.eager("EXPAND", 2, 4, |e, a| {
        let m = e.array_of(a[0].clone());
        let rows = whole(e, a, 1)?.unwrap_or(i64::from(m.rows));
        let cols = whole(e, a, 2)?.unwrap_or(i64::from(m.cols));
        if rows < i64::from(m.rows) || cols < i64::from(m.cols) || cols < 1 {
            return Err(DOMAIN);
        }
        let pad = opt(a, 3).cloned().unwrap_or_else(na);
        let (rows, cols) = (rows as usize, cols as usize);
        let mut data = Vec::with_capacity(rows * cols);
        for r in 0..rows {
            for c in 0..cols {
                data.push(if r < m.rows as usize && c < m.cols as usize {
                    m.get(r as u32, c as u32).clone()
                } else {
                    pad.clone()
                });
            }
        }
        Ok(array(rows, cols, data))
    })
    .array_args();

    r.eager("TOCOL", 1, 3, |e, a| to_vector(e, a, true))
        .array_args();
    r.eager("TOROW", 1, 3, |e, a| to_vector(e, a, false))
        .array_args();

    r.eager("WRAPROWS", 2, 3, |e, a| wrap(e, a, true))
        .array_args();
    r.eager("WRAPCOLS", 2, 3, |e, a| wrap(e, a, false))
        .array_args();

    r.eager("RANDARRAY", 0, 5, |e, a| {
        let rows = opt_num(e, a, 0, 1.0)?.trunc();
        let cols = opt_num(e, a, 1, 1.0)?.trunc();
        let (lo, hi) = (opt_num(e, a, 2, 0.0)?, opt_num(e, a, 3, 1.0)?);
        let whole = flag(e, a, 4)?;
        if rows < 1.0 || cols < 1.0 || lo > hi {
            return Err(DOMAIN);
        }
        let (rows, cols) = (rows as usize, cols as usize);
        let mut data = Vec::with_capacity(rows * cols);
        for _ in 0..rows * cols {
            let u = e.rand();
            data.push(Value::Number(if whole {
                let (l, h) = (lo.ceil(), hi.floor());
                if l > h {
                    return Err(DOMAIN);
                }
                l + (u * (h - l + 1.0)).floor()
            } else {
                lo + u * (hi - lo)
            }));
        }
        Ok(array(rows, cols, data))
    })
    .volatile();

    // ----- text -----
    r.eager("TEXTSPLIT", 2, 6, |e, a| {
        let text = e.text(&a[0])?;
        let cols = match opt(a, 1) {
            Some(v) => delimiters(e, v)?,
            None => Vec::new(),
        };
        let rows = match opt(a, 2) {
            Some(v) => delimiters(e, v)?,
            None => Vec::new(),
        };
        let ignore_empty = flag(e, a, 3)?;
        let ci = match_mode(e, a, 4)?;
        let pad = opt(a, 5).cloned().unwrap_or_else(na);
        if text.is_empty() {
            return Err(DOMAIN);
        }
        let mut table: Vec<Vec<String>> = Vec::new();
        for line in pieces(&text, &rows, ci) {
            let mut cells = pieces(&line, &cols, ci);
            if ignore_empty {
                cells.retain(|c| !c.is_empty());
                if cells.is_empty() {
                    continue;
                }
            }
            table.push(cells);
        }
        if table.is_empty() {
            return Err(ErrorKind::Na);
        }
        let width = table.iter().map(Vec::len).max().unwrap_or(1);
        let mut data = Vec::with_capacity(table.len() * width);
        for row in &table {
            for c in 0..width {
                data.push(row.get(c).map_or_else(|| pad.clone(), Value::text));
            }
        }
        Ok(array(table.len(), width, data))
    });
    r.eager("TEXTBEFORE", 2, 6, |e, a| text_around(e, a, true));
    r.eager("TEXTAFTER", 2, 6, |e, a| text_around(e, a, false));
    r.eager("ENCODEURL", 1, 1, |e, a| {
        let t = e.text(&a[0])?;
        if t.is_empty() {
            return Err(ErrorKind::Value);
        }
        Ok(Value::text(url_encode(&t)))
    });
    r.eager("ROT13", 1, 1, |e, a| {
        let t = e.text(&a[0])?;
        let rot = |c: char, base: u8| char::from((c as u8 - base + 13) % 26 + base);
        Ok(Value::text(
            t.chars()
                .map(|c| match c {
                    'a'..='z' => rot(c, b'a'),
                    'A'..='Z' => rot(c, b'A'),
                    _ => c,
                })
                .collect::<String>(),
        ))
    });
}

fn window(m: &Array, r0: usize, r1: usize, c0: usize, c1: usize) -> Value {
    let mut data = Vec::with_capacity((r1 - r0) * (c1 - c0));
    for r in r0..r1 {
        for c in c0..c1 {
            data.push(m.get(r as u32, c as u32).clone());
        }
    }
    array(r1 - r0, c1 - c0, data)
}

/// `CHOOSECOLS` / `CHOOSEROWS`: the numbers of the lanes to take, from the end when negative.
fn choose(e: &mut Engine, a: &[Value], cols: bool) -> Res<Value> {
    let m = e.array_of(a[0].clone());
    let len = i64::from(if cols { m.cols } else { m.rows });
    let mut which = Vec::new();
    for v in &a[1..] {
        for x in &e.array_of(v.clone()).data {
            let n = match x {
                Value::Number(n) => n.trunc() as i64,
                Value::Bool(b) => i64::from(*b),
                _ => return Err(DOMAIN),
            };
            if n == 0 || n.abs() > len {
                return Err(DOMAIN);
            }
            which.push(if n > 0 { n - 1 } else { len + n } as usize);
        }
    }
    Ok(if cols {
        pick_cols(&m, &which)
    } else {
        pick_rows(&m, &which)
    })
}

/// `HSTACK` / `VSTACK`: the arrays side by side (or one above the other), `#N/A` where one is
/// shorter.
fn stack(e: &mut Engine, a: &[Value], horizontal: bool) -> Res<Value> {
    let parts: Vec<Arc<Array>> = a.iter().map(|v| e.array_of(v.clone())).collect();
    let (rows, cols) = if horizontal {
        (
            parts.iter().map(|p| p.rows).max().unwrap_or(1) as usize,
            parts.iter().map(|p| p.cols as usize).sum(),
        )
    } else {
        (
            parts.iter().map(|p| p.rows as usize).sum(),
            parts.iter().map(|p| p.cols).max().unwrap_or(1) as usize,
        )
    };
    let mut data = vec![na(); rows * cols];
    let (mut r0, mut c0) = (0usize, 0usize);
    for p in &parts {
        for r in 0..p.rows as usize {
            for c in 0..p.cols as usize {
                data[(r0 + r) * cols + c0 + c] = p.get(r as u32, c as u32).clone();
            }
        }
        if horizontal {
            c0 += p.cols as usize;
        } else {
            r0 += p.rows as usize;
        }
    }
    Ok(array(rows, cols, data))
}

/// `TOCOL` / `TOROW`: the elements in one line, by rows or by columns, leaving out blanks and
/// errors as asked.
fn to_vector(e: &mut Engine, a: &[Value], column: bool) -> Res<Value> {
    if !matches!(a[0], Value::Ref(_) | Value::Array(_)) {
        return Err(DOMAIN);
    }
    let m = e.array_of(a[0].clone());
    let ignore = whole(e, a, 1)?.unwrap_or(0);
    if !(0..=3).contains(&ignore) {
        return Err(DOMAIN);
    }
    let by_col = flag(e, a, 2)?;
    let mut out = Vec::new();
    let mut take = |v: &Value| {
        let skip = (ignore & 1 != 0 && matches!(v, Value::Empty))
            || (ignore & 2 != 0 && matches!(v, Value::Error(_)));
        if !skip {
            out.push(v.clone());
        }
    };
    if by_col {
        for c in 0..m.cols {
            for r in 0..m.rows {
                take(m.get(r, c));
            }
        }
    } else {
        m.data.iter().for_each(&mut take);
    }
    if out.is_empty() {
        return Err(ErrorKind::Na);
    }
    let n = out.len();
    Ok(if column {
        array(n, 1, out)
    } else {
        array(1, n, out)
    })
}

/// `WRAPROWS` / `WRAPCOLS`: a vector cut into rows (or columns) of `wrap_count`.
fn wrap(e: &mut Engine, a: &[Value], by_rows: bool) -> Res<Value> {
    let m = e.array_of(a[0].clone());
    if m.rows != 1 && m.cols != 1 {
        return Err(DOMAIN);
    }
    let size = whole(e, a, 1)?.unwrap_or(0);
    if size < 1 {
        return Err(DOMAIN);
    }
    let size = size as usize;
    let pad = opt(a, 2).cloned().unwrap_or_else(na);
    let n = m.data.len();
    let lines = n.div_ceil(size);
    let (rows, cols) = if by_rows {
        (lines, size)
    } else {
        (size, lines)
    };
    let mut data = vec![pad; rows * cols];
    for (k, v) in m.data.iter().enumerate() {
        let (r, c) = if by_rows {
            (k / size, k % size)
        } else {
            (k % size, k / size)
        };
        data[r * cols + c] = v.clone();
    }
    Ok(array(rows, cols, data))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn delimiters_are_found_without_overlap_from_either_side() {
        let d = vec!["aa".to_string()];
        assert_eq!(forward("aaa", &d, false), vec![(0, 2)]);
        assert_eq!(backward("aaa", &d, false), vec![(1, 3)]);
        let d = vec!["-".to_string(), "b".to_string()];
        assert_eq!(forward("a-b-c", &d, false), vec![(1, 2), (2, 3), (3, 4)]);
        let d = vec!["A".to_string()];
        assert_eq!(forward("xAa", &d, true), vec![(1, 2), (2, 3)]);
    }

    #[test]
    fn url_encoding_keeps_letters_digits_hyphen_and_underscore() {
        assert_eq!(url_encode("a b&c/ñ-._~"), "a%20b%26c%2F%C3%B1-%2E_%7E");
    }
}
