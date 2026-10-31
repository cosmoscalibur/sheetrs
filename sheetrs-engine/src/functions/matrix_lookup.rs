//! Approximate lookups over an array that is not a range (an inline array or a computed one).
//!
//! LibreOffice searches those with code of their own, not the sorted-column search of a range, and
//! the positions it gives differ when the data are not sorted. These follow `ScMatch`,
//! `ScLookup` and `CalculateLookup` of `sc/source/core/tool/interpr1.cxx`.

use crate::value::{Array, Coll, Value};
use std::cmp::Ordering;

/// What LibreOffice calls a value in a matrix: numbers, booleans and errors (an error is a number
/// that is not finite).
fn is_value(v: &Value) -> bool {
    matches!(v, Value::Number(_) | Value::Bool(_) | Value::Error(_))
}

fn is_string_or_empty(v: &Value) -> bool {
    matches!(v, Value::Text(_) | Value::Empty)
}

fn number_of(v: &Value) -> f64 {
    match v {
        Value::Number(n) => *n,
        Value::Bool(b) => f64::from(u8::from(*b)),
        _ => f64::NAN,
    }
}

fn string_of(v: &Value) -> &str {
    match v {
        Value::Text(t) => t,
        _ => "",
    }
}

/// `-1`, `0` or `1` when the matrix value is smaller than, equal to or larger than the searched
/// one (`lcl_CompareMatrix2Query`). Numbers sort before text, errors after both; an empty value is
/// the smallest, or the largest when `empty_is_less` is off.
fn cmp_query(v: &Value, needle: &Value, coll: &Coll, empty_is_less: bool) -> i32 {
    if matches!(v, Value::Empty) {
        return if empty_is_less { -1 } else { 1 };
    }
    let by_string = matches!(needle, Value::Text(_));
    if is_value(v) {
        let n = number_of(v);
        if !n.is_finite() {
            return 1;
        }
        if by_string {
            return -1;
        }
        let q = number_of(needle);
        return if n == q {
            0
        } else if n < q {
            -1
        } else {
            1
        };
    }
    if !by_string {
        return 1;
    }
    match coll.compare(string_of(v), string_of(needle)) {
        Ordering::Less => -1,
        Ordering::Equal => 0,
        Ordering::Greater => 1,
    }
}

/// The last of the equal values that follow `i` (`lcl_GetLastMatch`).
fn last_match(lane: &[Value], mut i: usize) -> usize {
    let n = lane.len();
    match &lane[i] {
        v if is_value(v) => {
            let x = number_of(v);
            while i + 1 < n && is_value(&lane[i + 1]) && x == number_of(&lane[i + 1]) {
                i += 1;
            }
        }
        Value::Empty => {
            while i + 1 < n && matches!(lane[i + 1], Value::Empty) {
                i += 1;
            }
        }
        Value::Text(s) => {
            while i + 1 < n && matches!(&lane[i + 1], Value::Text(t) if t == s) {
                i += 1;
            }
        }
        _ => {}
    }
    i
}

/// The lane of an array that is searched: its first column when it is at least as tall as wide,
/// else its first row.
pub(super) fn lane(m: &Array) -> (Vec<Value>, bool) {
    let vertical = m.rows >= m.cols;
    let lane = if vertical {
        (0..m.rows).map(|r| m.get(r, 0).clone()).collect()
    } else {
        (0..m.cols).map(|c| m.get(0, c).clone()).collect()
    };
    (lane, vertical)
}

/// `MATCH` of type 1 (`ascending`) or -1 over the elements of an array: 0-based position.
pub(super) fn match_position(
    needle: &Value,
    lane: &[Value],
    ascending: bool,
    coll: &Coll,
) -> Option<usize> {
    let n = lane.len();
    let cmp = |i: usize| cmp_query(&lane[i], needle, coll, true);
    let (mut first, mut last, mut hit) = (0usize, n - 1, 0usize);
    loop {
        let len = last - first;
        if len == 0 {
            break;
        }
        let mid = first + len / 2;
        let c = cmp(mid);
        if c == 0 {
            // An exact hit is final, of whatever type it is.
            return Some(last_match(lane, mid));
        }
        if len == 1 {
            hit = if (c < 0) == ascending { last } else { first };
            break;
        }
        if (c < 0) == ascending {
            first = mid;
        } else {
            last = mid;
        }
    }
    if hit == n - 1 {
        let c = cmp(hit);
        if (ascending && c <= 0) || (!ascending && c >= 0) {
            return Some(hit);
        }
    }
    // `hit` counts from 1 here: the position before it is the answer.
    if hit == 0 {
        return None;
    }
    let by_string = matches!(needle, Value::Text(_));
    if by_string == is_value(&lane[hit - 1]) {
        return None;
    }
    Some(hit - 1)
}

/// `LOOKUP` over an array: 0-based position in the lane. `numeric` says that the whole array holds
/// only numbers, booleans and errors, in which case the errors are left out of the search.
pub(super) fn lookup_position(
    needle: &Value,
    full: &[Value],
    numeric: bool,
    coll: &Coll,
) -> Option<usize> {
    let kept: Vec<usize> = if numeric {
        (0..full.len())
            .filter(|&i| number_of(&full[i]).is_finite())
            .collect()
    } else {
        (0..full.len()).collect()
    };
    if kept.is_empty() {
        return None;
    }
    let lane: Vec<Value> = kept.iter().map(|&i| full[i].clone()).collect();
    let n = lane.len();
    let cmp = |i: usize| cmp_query(&lane[i], needle, coll, false);
    let (mut first, mut last) = (0usize, n - 1);
    let mut delta: isize = -1;
    let mut found = false;
    loop {
        let len = last - first;
        if len == 0 {
            break;
        }
        let mid = first + len / 2;
        let c = cmp(mid);
        if c == 0 {
            delta = last_match(&lane, mid).cast_signed();
            found = true;
            break;
        }
        if len == 1 {
            delta = if c < 0 {
                last.cast_signed() - 1
            } else {
                first.cast_signed() - 1
            };
            found = delta >= 0;
            break;
        }
        if c < 0 {
            first = mid;
        } else {
            last = mid;
        }
    }
    if delta == n.cast_signed() - 2 {
        if cmp((delta + 1) as usize) <= 0 {
            delta += 1;
            found = true;
        }
    } else if delta > 0 {
        found = true;
    }
    if !found {
        return None;
    }
    let pos = kept[delta as usize];
    // Numbers against text (or the reverse) are #N/A, as in Excel.
    let by_string = matches!(needle, Value::Text(_));
    if by_string == is_value(&full[pos]) {
        return None;
    }
    Some(pos)
}

/// `VLOOKUP`/`HLOOKUP` over an array: 0-based position of the key in the first column (row).
/// `sorted` is the approximate search; otherwise the first equal key is taken, without wildcards.
pub(super) fn table_position(
    needle: &Value,
    keys: &[Value],
    sorted: bool,
    coll: &Coll,
    equal_text: &dyn Fn(&str, &str) -> bool,
) -> Option<usize> {
    let n = keys.len();
    let mut delta = None;
    match needle {
        Value::Text(p) => {
            if sorted {
                let mut i = 0;
                while i < n {
                    if is_string_or_empty(&keys[i]) {
                        if coll.compare(string_of(&keys[i]), p) != Ordering::Greater {
                            delta = Some(i);
                        } else if i > 0 {
                            break;
                        }
                    } else {
                        delta = Some(i);
                    }
                    i += 1;
                }
            } else {
                delta = keys
                    .iter()
                    .position(|k| is_string_or_empty(k) && equal_text(string_of(k), p));
            }
        }
        _ => {
            let x = number_of(needle);
            for (i, k) in keys.iter().enumerate() {
                if is_string_or_empty(k) {
                    continue;
                }
                let v = number_of(k);
                if sorted {
                    if v <= x {
                        delta = Some(i);
                    } else {
                        break;
                    }
                } else if v == x {
                    delta = Some(i);
                    break;
                }
            }
        }
    }
    delta
}

/// How `XLOOKUP` compares: equal, or the nearest smaller (`LessEq`) or larger (`GreaterEq`).
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum Op {
    Equal,
    LessEq,
    GreaterEq,
}

/// `lcl_Compare2MatrixCells`: two matrix values against each other.
fn cmp_cells(a: &Value, b: &Value, coll: &Coll) -> i32 {
    if matches!(a, Value::Empty) {
        return if matches!(b, Value::Empty) { 0 } else { -1 };
    }
    if matches!(b, Value::Empty) {
        return 1;
    }
    let by_string = matches!(b, Value::Text(_));
    if is_value(a) {
        let x = number_of(a);
        if !x.is_finite() {
            return 1;
        }
        if by_string {
            return -1;
        }
        let y = number_of(b);
        return if x == y {
            0
        } else if x < y {
            -1
        } else {
            1
        };
    }
    if !by_string {
        return 1;
    }
    match coll.compare(string_of(a), string_of(b)) {
        Ordering::Less => -1,
        Ordering::Equal => 0,
        Ordering::Greater => 1,
    }
}

/// How `XLOOKUP` walks an array (`SearchMatrixForValue`).
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum Walk {
    Forward,
    Reverse,
    /// The binary searches of search modes 2 and -2, which LibreOffice runs alike.
    Binary,
}

/// `XLOOKUP` of a value that is not empty over the elements of an array that is not a range: the
/// 0-based position found.
pub(super) fn xlookup_position(
    needle: &Value,
    lane: &[Value],
    op: Op,
    walk: Walk,
    coll: &Coll,
) -> Option<usize> {
    let n = lane.len();
    let cmp = |i: usize| cmp_query(&lane[i], needle, coll, true);
    let mut hit = 0usize; // 1-based; 0 is none
    let mut best: Option<usize> = None;
    let better = |best: &mut Option<usize>, i: usize, want: i32| match *best {
        None => *best = Some(i),
        Some(b) if cmp_cells(&lane[i], &lane[b], coll) == want => *best = Some(i),
        _ => {}
    };
    match walk {
        Walk::Forward | Walk::Reverse => {
            let reverse = walk == Walk::Reverse;
            // The walk down skips the last element when it looks for the nearest value.
            let order: Vec<usize> = match (reverse, op) {
                (false, _) => (0..n).collect(),
                (true, Op::Equal) => (0..n).rev().collect(),
                (true, _) => (0..n.saturating_sub(1)).rev().collect(),
            };
            for i in order {
                let c = cmp(i);
                if c == 0 {
                    hit = i + 1;
                    break;
                }
                match op {
                    Op::LessEq if c == -1 => better(&mut best, i, 1),
                    Op::GreaterEq if c == 1 => better(&mut best, i, -1),
                    _ => {}
                }
            }
        }
        Walk::Binary => {
            let asc = op == Op::LessEq;
            let (mut first, mut last) = (0usize, n - 1);
            let mut h = 0usize;
            loop {
                let len = last - first;
                if len == 0 {
                    break;
                }
                let mid = first + len / 2;
                let c = cmp(mid);
                if c == 0 {
                    return Some(last_match(lane, mid));
                }
                if len == 1 {
                    h = if (c < 0) == asc { last } else { first };
                    break;
                }
                if (c < 0) == asc {
                    first = mid;
                } else {
                    last = mid;
                }
            }
            if h == n - 1 {
                let c = cmp(h);
                if (asc && c <= 0) || (!asc && c >= 0) {
                    return Some(h);
                }
            }
            hit = h;
        }
    }
    if hit > 0 {
        let by_string = matches!(needle, Value::Text(_));
        if by_string == is_value(&lane[hit - 1]) {
            return None;
        }
        return Some(hit - 1);
    }
    best
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::value::Collators;

    fn nums(v: &[f64]) -> Vec<Value> {
        v.iter().map(|n| Value::Number(*n)).collect()
    }

    #[test]
    fn match_walks_the_array_as_libreoffice_does() {
        let c = Collators::new("en");
        let col = c.get(false);
        // Not sorted: the answers LibreOffice 26.8 gave.
        let p = |a: &[f64], x: f64, asc: bool| {
            match_position(&Value::Number(x), &nums(a), asc, col).map(|i| i + 1)
        };
        assert_eq!(p(&[4.0, 2.0, 6.0, 8.0, 10.0], 4.0, true), Some(2));
        assert_eq!(p(&[3.0, 2.0, 3.0, 3.0], 3.0, true), Some(4));
        assert_eq!(p(&[1.0, 8.0], 3.0, false), None);
        assert_eq!(p(&[6.0, 9.0], 6.0, false), Some(1));
        assert_eq!(p(&[9.0, 5.0, 2.0], 5.0, true), Some(2));
    }
}
