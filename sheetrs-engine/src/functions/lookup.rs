//! Lookup and reference functions.

use super::matrix_lookup;
use super::stats::values_equal;
use super::{Reg, num, opt, opt_num};
use crate::addr::{self, MAX_COLS, MAX_ROWS, Rect};
use crate::engine::{Engine, Res};
use crate::matching::Matcher;
use crate::options::{EngineOptions, Search};
use crate::value::Coll;
use crate::value::{Array, ErrorKind, Value, compare};
use std::cmp::Ordering;
use std::fmt::Write as _;
use std::sync::Arc;

pub(super) fn register(r: &mut Reg) {
    r.eager("MATCH", 2, 3, |e, a| {
        let needle = e.scalar(a[0].clone());
        let hay = e.array_of(a[1].clone());
        let kind = opt_num(e, a, 2, 1.0)?.trunc();
        if !(-1.0..=1.0).contains(&kind) {
            return Err(ErrorKind::Domain);
        }
        if hay.rows > 1 && hay.cols > 1 {
            return Err(ErrorKind::Domain);
        }
        // An array that is not a range has its own search for the types 1 and -1.
        if !matches!(a[1], Value::Ref(_)) && kind != 0.0 && !matches!(needle, Value::Empty) {
            return matrix_lookup::match_position(
                &needle,
                &hay.data,
                kind > 0.0,
                e.collator(false),
            )
            .map_or(Err(ErrorKind::Na), |p| num((p + 1) as f64));
        }
        let found = position(&needle, &hay.data, kind, &e.opts, e.collator(false));
        // In a range, a number is also found by the text it shows: `MATCH("5", A1:A8, 0)` finds 5,
        // and `MATCH("*", A1:A8, 0)` finds the first number.
        let found = found.or_else(|| match (&needle, &a[1]) {
            (Value::Text(t), Value::Ref(_)) if kind == 0.0 => {
                find_shown(e, t, hay.data.len(), |i| &hay.data[i])
            }
            _ => None,
        });
        found.map_or(Err(ErrorKind::Na), |p| num((p + 1) as f64))
    })
    .array_args();
    r.eager("INDEX", 2, 4, index).array_args();
    r.eager("VLOOKUP", 3, 4, |e, a| table_lookup(e, a, false));
    r.eager("HLOOKUP", 3, 4, |e, a| table_lookup(e, a, true));
    r.eager("XLOOKUP", 3, 6, xlookup).array_args();
    r.eager("XMATCH", 2, 4, xmatch).array_args();
    r.eager("OFFSET", 3, 5, offset).dynamic();
    r.eager("INDIRECT", 1, 2, indirect).dynamic();
    r.eager("_RANGE", 2, 2, |_, a| match (&a[0], &a[1]) {
        (Value::Ref(x), Value::Ref(y)) if x.sheet == y.sheet => Ok(Value::Ref(Rect::new(
            x.sheet,
            x.r0.min(y.r0),
            x.c0.min(y.c0),
            x.r1.max(y.r1),
            x.c1.max(y.c1),
        ))),
        _ => Err(ErrorKind::Ref),
    })
    .dynamic();
    r.eager("ROW", 0, 1, |e, a| match a.first() {
        None => num(f64::from(e.at.row + 1)),
        Some(Value::Ref(r)) => num(f64::from(r.r0 + 1)),
        Some(_) => Err(ErrorKind::Value),
    });
    r.eager("COLUMN", 0, 1, |e, a| match a.first() {
        None => num(f64::from(e.at.col + 1)),
        Some(Value::Ref(r)) => num(f64::from(r.c0 + 1)),
        Some(_) => Err(ErrorKind::Value),
    });
    r.eager("ROWS", 1, 1, |_, a| match &a[0] {
        Value::Error(k) => Err(*k),
        Value::Ref(r) => num(f64::from(r.rows())),
        Value::Array(m) => num(f64::from(m.rows)),
        _ => num(1.0),
    });
    r.eager("COLUMNS", 1, 1, |_, a| match &a[0] {
        Value::Error(k) => Err(*k),
        Value::Ref(r) => num(f64::from(r.cols())),
        Value::Array(m) => num(f64::from(m.cols)),
        _ => num(1.0),
    });
    r.eager("ADDRESS", 2, 5, address);
    r.eager("AREAS", 1, 1, |_, a| match &a[0] {
        Value::Ref(_) => num(1.0),
        Value::Error(k) => Err(*k),
        _ => Err(ErrorKind::Domain),
    });
    r.eager("SHEET", 0, 1, |e, a| match a.first() {
        None => num(f64::from(e.at.sheet) + 1.0),
        Some(Value::Ref(r)) => num(f64::from(r.sheet) + 1.0),
        Some(Value::Error(k)) => Err(*k),
        Some(Value::Text(t)) => {
            let found = e
                .sheet_names()
                .iter()
                .position(|s| s.eq_ignore_ascii_case(t));
            found.map_or(Err(ErrorKind::Domain), |i| num((i + 1) as f64))
        }
        Some(_) => Err(ErrorKind::Domain),
    });
    r.eager("SHEETS", 0, 1, |e, a| match a.first() {
        None => num(e.sheet_names().len() as f64),
        Some(Value::Ref(_)) => num(1.0),
        Some(Value::Error(k)) => Err(*k),
        Some(_) => Err(ErrorKind::Domain),
    });
    // The text of the link, or its address when it has none; a blank is 0 and a boolean a number.
    r.eager("HYPERLINK", 1, 2, |e, a| {
        let pick = if a.len() > 1 { &a[1] } else { &a[0] };
        let url = e.scalar(a[0].clone());
        if let Value::Error(k) = url {
            return Err(k);
        }
        Ok(match e.scalar(pick.clone()) {
            Value::Error(k) => return Err(k),
            Value::Empty => Value::Number(0.0),
            Value::Bool(b) => Value::Number(f64::from(u8::from(b))),
            v => v,
        })
    });
}

/// Position (0-based) of `needle` in `hay` for a MATCH `kind` (0 exact,
/// 1 largest ≤ in ascending data, -1 smallest ≥ in descending data).
pub(super) fn position(
    needle: &Value,
    hay: &[Value],
    kind: f64,
    opts: &EngineOptions,
    coll: &Coll,
) -> Option<usize> {
    position_with(needle, hay.len(), |i| &hay[i], kind, opts, coll)
}

/// [`position`] over `len` values reached through `at`, so that a column of a
/// table is searched where it lies instead of being copied out first.
pub(super) fn position_with<'a>(
    needle: &Value,
    len: usize,
    at: impl Fn(usize) -> &'a Value,
    kind: f64,
    opts: &EngineOptions,
    coll: &Coll,
) -> Option<usize> {
    // Nothing is found for an empty cell as the searched value.
    if matches!(needle, Value::Empty) {
        return None;
    }
    let search = &opts.search;
    if kind == 0.0 {
        if let Value::Text(p) = needle {
            let m = Matcher::new(p, search, search.whole_cell);
            return (0..len).find(|&i| matches!(at(i), Value::Text(t) if m.is_match(t)));
        }
        return (0..len).find(|&i| values_equal(needle, at(i), coll));
    }
    // Booleans are numbers, as in LibreOffice.
    let same_class = |v: &Value| {
        (matches!(needle, Value::Number(_) | Value::Bool(_))
            && matches!(v, Value::Number(_) | Value::Bool(_)))
            || matches!((needle, v), (Value::Text(_), Value::Text(_)))
    };
    let mut best = None;
    for i in 0..len {
        let v = at(i);
        if !same_class(v) {
            continue;
        }
        let ord = compare(v, needle, coll);
        let ok = if kind > 0.0 {
            ord != Ordering::Greater
        } else {
            ord != Ordering::Less
        };
        if ok {
            best = Some(i);
        } else {
            break;
        }
    }
    best
}

fn index(e: &mut Engine, a: &[Value]) -> Res<Value> {
    // The index of an error is that error, wherever it points.
    if let Value::Error(k) = &a[0] {
        return Err(*k);
    }
    let row = opt(a, 1).map(|v| e.num(v)).transpose()?;
    let col = opt(a, 2).map(|v| e.num(v)).transpose()?;
    let (rows, cols) = match &a[0] {
        Value::Ref(r) => (r.rows(), r.cols()),
        Value::Array(m) => (m.rows, m.cols),
        _ => (1, 1),
    };
    let (row, col) = match (row, col) {
        (Some(r), None) if cols == 1 => (r, 1.0),
        (Some(r), None) if rows == 1 => (1.0, r),
        (r, c) => (r.unwrap_or(0.0), c.unwrap_or(0.0)),
    };
    if row < 0.0 || col < 0.0 || row > f64::from(rows) || col > f64::from(cols) {
        return Err(ErrorKind::Ref);
    }
    let (row, col) = (row as u32, col as u32);
    let (r0, r1) = if row == 0 {
        (0, rows - 1)
    } else {
        (row - 1, row - 1)
    };
    let (c0, c1) = if col == 0 {
        (0, cols - 1)
    } else {
        (col - 1, col - 1)
    };
    match &a[0] {
        Value::Ref(r) => Ok(Value::Ref(Rect::new(
            r.sheet,
            r.r0 + r0,
            r.c0 + c0,
            r.r0 + r1,
            r.c0 + c1,
        ))),
        other => {
            let m = e.array_of(other.clone());
            let (h, w) = (r1 - r0 + 1, c1 - c0 + 1);
            let mut data = Vec::with_capacity((h * w) as usize);
            for i in r0..=r1 {
                for j in c0..=c1 {
                    // A blank that went through an array function is an empty text, as there.
                    data.push(match m.get(i, j) {
                        Value::Empty => Value::text(""),
                        v => v.clone(),
                    });
                }
            }
            Ok(if data.len() == 1 {
                data.remove(0)
            } else {
                Value::Array(Arc::new(Array::new(h, w, data)))
            })
        }
    }
}

fn table_lookup(e: &mut Engine, a: &[Value], horizontal: bool) -> Res<Value> {
    let needle = e.scalar(a[0].clone());
    if let Value::Error(k) = needle {
        return Err(k);
    }
    let t = e.array_of(a[1].clone());
    let idx = e.num(&a[2])?.trunc();
    let approx = match opt(a, 3) {
        Some(v) => e.boolean(v)?,
        None => !matches!(a.get(3), Some(Value::Empty)),
    };
    let (lanes, depth) = if horizontal {
        (t.cols, t.rows)
    } else {
        (t.rows, t.cols)
    };
    if idx < 1.0 {
        return Err(ErrorKind::Value);
    }
    if idx as u32 > depth {
        return Err(ErrorKind::Ref);
    }
    // An array that is not a range is searched as LibreOffice does.
    if let Value::Array(m) = &a[1] {
        let keys: Vec<Value> = if horizontal {
            (0..m.cols).map(|c| m.get(0, c).clone()).collect()
        } else {
            (0..m.rows).map(|r| m.get(r, 0).clone()).collect()
        };
        let needle = match needle {
            Value::Empty => Value::text(""),
            n => n,
        };
        let coll = e.collator(false);
        let same = |x: &str, y: &str| values_equal(&Value::text(x), &Value::text(y), coll);
        let pos = matrix_lookup::table_position(&needle, &keys, approx, coll, &same)
            .ok_or(ErrorKind::Na)?;
        // Text sought, a number found (or the reverse for an exact key) is #N/A.
        if matches!(needle, Value::Text(_))
            && matches!(
                keys[pos],
                Value::Number(_) | Value::Bool(_) | Value::Error(_)
            )
        {
            return Err(ErrorKind::Na);
        }
        let k = idx as u32 - 1;
        let v = if horizontal {
            m.get(k, pos as u32)
        } else {
            m.get(pos as u32, k)
        };
        return Ok(match v {
            Value::Empty => Value::text(""),
            v => v.clone(),
        });
    }
    let key = |i: usize| {
        if horizontal {
            t.get(0, i as u32)
        } else {
            t.get(i as u32, 0)
        }
    };
    // An exact search of a big range, repeated for many keys, goes through a hash
    // index of its key lane (the first match, as the scan finds it).
    let indexed = (!approx)
        .then(|| match (&a[1], &needle) {
            (Value::Ref(r), Value::Text(p)) => {
                let m = Matcher::new(p, &e.opts.search, e.opts.search.whole_cell);
                let folded = m.plain_whole()?.to_string();
                let index = e.lookup_index(*r, horizontal)?;
                Some(index.find_text(&folded))
            }
            (Value::Ref(r), Value::Number(n)) if n.is_finite() => {
                let index = e.lookup_index(*r, horizontal)?;
                Some(index.find_number(*n))
            }
            _ => None,
        })
        .flatten();
    let found = match indexed {
        Some(hit) => hit,
        None => position_with(
            &needle,
            lanes as usize,
            key,
            if approx { 1.0 } else { 0.0 },
            &e.opts,
            e.collator(false),
        ),
    };
    let found = found.or_else(|| match (&a[1], &needle) {
        (Value::Ref(_), Value::Text(p)) if !approx => find_shown(e, p, lanes as usize, key),
        _ => None,
    });
    let pos = found.ok_or(ErrorKind::Na)? as u32;
    let k = idx as u32 - 1;
    // Over a range, the answer is the cell itself, so that an empty result keeps
    // being one (LibreOffice hands the cell's content over, not a copy of the value).
    if let Value::Ref(r) = &a[1] {
        let (dr, dc) = if horizontal { (k, pos) } else { (pos, k) };
        return Ok(Value::Ref(Rect::single(crate::addr::CellRef::new(
            r.sheet,
            r.r0 + dr,
            r.c0 + dc,
        ))));
    }
    Ok(if horizontal {
        t.get(k, pos).clone()
    } else {
        t.get(pos, k).clone()
    })
}

fn xlookup(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let needle = e.scalar(a[0].clone());
    let hay = e.array_of(a[1].clone());
    let mode = opt_num(e, a, 4, 0.0)?;
    let search_mode = opt_num(e, a, 5, 1.0)?;
    let Some(i) = xlookup_find(e, &needle, &a[1], &hay, mode, search_mode)? else {
        return match a.get(3) {
            Some(v) if !matches!(v, Value::Empty) => Ok(v.clone()),
            _ => Err(ErrorKind::Na),
        };
    };
    xlookup_result(e, a, &hay, i)
}

/// `XMATCH(lookup, array, [match_mode], [search_mode])`: the position `XLOOKUP` finds, from 1.
fn xmatch(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let needle = e.scalar(a[0].clone());
    let hay = e.array_of(a[1].clone());
    let mode = opt_num(e, a, 2, 0.0)?;
    let search_mode = opt_num(e, a, 3, 1.0)?;
    if hay.rows > 1 && hay.cols > 1 {
        return Err(ErrorKind::Domain);
    }
    match xlookup_find(e, &needle, &a[1], &hay, mode, search_mode)? {
        Some(i) => num((i + 1) as f64),
        None => Err(ErrorKind::Na),
    }
}

/// The text a number cell shows (General format), which a range search compares with a text
/// criterion.
fn shown_number(e: &Engine, v: &Value) -> Option<String> {
    match v {
        Value::Number(n) => crate::coerce::to_text(&Value::Number(*n), &e.opts.locale).ok(),
        _ => None,
    }
}

/// The first of `len` values (reached through `at`) that is a number whose shown text matches the
/// text criterion `p`. A plain criterion that is not a number cannot match one, so the scan is
/// skipped for it.
fn find_shown<'a>(
    e: &Engine,
    p: &str,
    len: usize,
    at: impl Fn(usize) -> &'a Value,
) -> Option<usize> {
    let m = Matcher::new(p, &e.opts.search, e.opts.search.whole_cell);
    if m.plain_whole().is_some_and(|s| s.parse::<f64>().is_err()) {
        return None;
    }
    (0..len).find(|&i| shown_number(e, at(i)).is_some_and(|s| m.is_match(&s)))
}

/// The index in `hay` that `XLOOKUP` (and `XMATCH`) find for `needle`, if any.
fn xlookup_find(
    e: &mut Engine,
    needle: &Value,
    source: &Value,
    hay: &Array,
    mode: f64,
    search_mode: f64,
) -> Res<Option<usize>> {
    if let Value::Error(k) = needle {
        return Err(*k);
    }
    let search_mode = search_mode.trunc() as i32;
    if search_mode == 0 || !(-2..=2).contains(&search_mode) || !(-1.0..=3.0).contains(&mode.trunc())
    {
        return Err(ErrorKind::Domain);
    }
    let reverse = search_mode < 0;
    let binary = search_mode.abs() == 2;
    if binary && matches!(source, Value::Ref(_)) {
        crate::notes::note(crate::notes::XLOOKUP_BINARY);
    }
    if binary && mode.trunc() >= 2.0 {
        return Err(ErrorKind::Value);
    }
    let mut order: Vec<usize> = (0..hay.data.len()).collect();
    if reverse {
        order.reverse();
    }
    // Over an array that is not a range, the nearest-value modes and the binary searches have code
    // of their own in LibreOffice.
    let on_array = !matches!(source, Value::Ref(_)) && !matches!(needle, Value::Empty);
    Ok(
        if on_array && (matches!(mode as i32, -1 | 1) || (mode as i32 == 0 && binary)) {
            if hay.rows > 1 && hay.cols > 1 {
                return Err(ErrorKind::Domain);
            }
            let op = match mode as i32 {
                -1 => matrix_lookup::Op::LessEq,
                1 => matrix_lookup::Op::GreaterEq,
                _ => matrix_lookup::Op::Equal,
            };
            let walk = match search_mode {
                1 => matrix_lookup::Walk::Forward,
                -1 => matrix_lookup::Walk::Reverse,
                _ => matrix_lookup::Walk::Binary,
            };
            matrix_lookup::xlookup_position(needle, &hay.data, op, walk, e.collator(false))
        } else {
            range_xlookup(
                e,
                needle,
                hay,
                mode as i32,
                &order,
                matches!(source, Value::Ref(_)),
            )
        },
    )
}

/// The position `XLOOKUP` finds in a range (or in an array, for the exact modes), looking at the
/// elements in `order`.
fn range_xlookup(
    e: &mut Engine,
    needle: &Value,
    hay: &Array,
    mode: i32,
    order: &[usize],
    from_range: bool,
) -> Option<usize> {
    let pick = |i: usize| hay.data[i].clone();
    // In a range, a number is also found by the text it shows: `"5"` finds 5.
    let shown = |e: &Engine, v: &Value| if from_range { shown_number(e, v) } else { None };
    match mode {
        0 => order.iter().copied().find(|&i| {
            values_equal(needle, &hay.data[i], e.collator(false))
                || matches!((needle, shown(e, &hay.data[i])), (Value::Text(t), Some(s)) if **t == *s)
        }),
        2 | 3 => match &needle {
            Value::Text(p) => {
                // The explicit wildcard and regular expression modes of XLOOKUP, whatever the
                // document says.
                let search = Search {
                    mode: if mode == 2 {
                        crate::options::SearchMode::Wildcards
                    } else {
                        crate::options::SearchMode::Regex
                    },
                    whole_cell: true,
                    ..e.opts.search
                };
                let wild = Matcher::new(p, &search, true);
                order.iter().copied().find(|&i| {
                    matches!(&hay.data[i], Value::Text(t) if wild.is_match(t))
                        || shown(e, &hay.data[i]).is_some_and(|s| wild.is_match(&s))
                })
            }
            _ => order
                .iter()
                .copied()
                .find(|&i| values_equal(needle, &hay.data[i], e.collator(false))),
        },
        m @ (-1 | 1) => {
            // Exact, else nearest smaller (-1) / larger (1).
            let mut best: Option<usize> = None;
            for &i in order {
                let v = pick(i);
                if !matches!(
                    (&needle, &v),
                    (
                        Value::Number(_) | Value::Bool(_),
                        Value::Number(_) | Value::Bool(_)
                    ) | (Value::Text(_), Value::Text(_))
                ) {
                    continue;
                }
                match compare(&v, needle, e.collator(false)) {
                    Ordering::Equal => {
                        best = Some(i);
                        break;
                    }
                    o if (m == -1 && o == Ordering::Less) || (m == 1 && o == Ordering::Greater) => {
                        let better = best.is_none_or(|b| {
                            let ob = compare(&v, &hay.data[b], e.collator(false));
                            if m == -1 {
                                ob == Ordering::Greater
                            } else {
                                ob == Ordering::Less
                            }
                        });
                        if better {
                            best = Some(i);
                        }
                    }
                    _ => {}
                }
            }
            best
        }
        _ => None,
    }
}

fn xlookup_result(e: &mut Engine, a: &[Value], hay: &Array, i: usize) -> Res<Value> {
    // The match index applies along the orientation of the lookup array.
    let vertical = hay.cols == 1;
    match &a[2] {
        Value::Ref(r) => {
            let i = i as u32;
            Ok(Value::Ref(if vertical {
                Rect::new(r.sheet, r.r0 + i, r.c0, r.r0 + i, r.c1)
            } else {
                Rect::new(r.sheet, r.r0, r.c0 + i, r.r1, r.c0 + i)
            }))
        }
        other => {
            let m = e.array_of(other.clone());
            let i = i as u32;
            if vertical && m.cols == 1 || !vertical && m.rows == 1 {
                Ok(m.data.get(i as usize).cloned().ok_or(ErrorKind::Ref)?)
            } else if vertical {
                let data: Vec<Value> = (0..m.cols).map(|j| m.get(i, j).clone()).collect();
                Ok(Value::Array(Arc::new(Array::new(1, m.cols, data))))
            } else {
                let data: Vec<Value> = (0..m.rows).map(|j| m.get(j, i).clone()).collect();
                Ok(Value::Array(Arc::new(Array::new(m.rows, 1, data))))
            }
        }
    }
}

fn offset(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let Value::Ref(base) = &a[0] else {
        return Err(ErrorKind::Value);
    };
    let dr = e.num(&a[1])?.trunc() as i64;
    let dc = e.num(&a[2])?.trunc() as i64;
    let h = opt_num(e, a, 3, f64::from(base.rows()))?.trunc() as i64;
    let w = opt_num(e, a, 4, f64::from(base.cols()))?.trunc() as i64;
    let (r0, c0) = (i64::from(base.r0) + dr, i64::from(base.c0) + dc);
    let (r1, c1) = (r0 + h - 1, c0 + w - 1);
    if h < 1 || w < 1 || r0 < 0 || c0 < 0 || r1 >= i64::from(MAX_ROWS) || c1 >= i64::from(MAX_COLS)
    {
        return Err(ErrorKind::Ref);
    }
    Ok(Value::Ref(Rect::new(
        base.sheet, r0 as u32, c0 as u32, r1 as u32, c1 as u32,
    )))
}

fn indirect(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let text = e.text(&a[0])?;
    if let Some(v) = opt(a, 1)
        && !e.boolean(v)?
    {
        return Err(ErrorKind::Ref); // R1C1 notation is not supported.
    }
    let (sheet, tail) = addr::string_ref(&text, e.opts.string_ref_excel).ok_or(ErrorKind::Ref)?;
    let sheet = match sheet {
        Some(n) => e.sheet_id(&n).ok_or(ErrorKind::Ref)?,
        None => e.at.sheet,
    };
    addr::parse_area(&tail)
        .map(|ar| Value::Ref(addr::bind(ar, sheet)))
        .ok_or(ErrorKind::Ref)
}

fn address(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let (r, c) = (e.num(&a[0])?.trunc(), e.num(&a[1])?.trunc());
    if r < 1.0 || c < 1.0 || r > f64::from(MAX_ROWS) || c > f64::from(MAX_COLS) {
        return Err(ErrorKind::Value);
    }
    let abs = opt_num(e, a, 2, 1.0)? as i32;
    let a1 = match opt(a, 3) {
        Some(v) => e.boolean(v)?,
        None => true,
    };
    let (row_abs, col_abs) = match abs {
        1 => (true, true),
        2 => (true, false),
        3 => (false, true),
        4 => (false, false),
        _ => return Err(ErrorKind::Value),
    };
    let mut s = String::new();
    if let Some(v) = opt(a, 4) {
        let name = e.text(v)?;
        if name.chars().all(|c| c.is_alphanumeric() || c == '_') {
            s.push_str(&name);
        } else {
            let _ = write!(s, "'{}'", name.replace('\'', "''"));
        }
        // The sheet separator follows the document's reference syntax (not in R1C1).
        s.push(if a1 && !e.opts.string_ref_excel {
            '.'
        } else {
            '!'
        });
    }
    let (r, c) = (r as u32, c as u32);
    if a1 {
        let _ = write!(
            s,
            "{}{}{}{}",
            if col_abs { "$" } else { "" },
            addr::col_name(c - 1),
            if row_abs { "$" } else { "" },
            r
        );
    } else {
        let rr = if row_abs {
            format!("R{r}")
        } else {
            format!("R[{}]", i64::from(r) - i64::from(e.at.row + 1))
        };
        let cc = if col_abs {
            format!("C{c}")
        } else {
            format!("C[{}]", i64::from(c) - i64::from(e.at.col + 1))
        };
        s.push_str(&rr);
        s.push_str(&cc);
    }
    Ok(Value::text(s))
}
