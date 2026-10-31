//! The database functions `DSUM`, `DCOUNT`, `DCOUNTA`, `DAVERAGE`, `DGET`, `DMAX`, `DMIN`,
//! `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DVAR` and `DVARP`: a table whose first row names its fields, a
//! field (a name or a position) and a block of criteria (a row of names and one row per
//! alternative, whose cells all have to hold). What LibreOffice 26.8 answers for the arguments
//! outside that shape is what `tests/assets/semantics.*` holds.

use super::stats::Criterion;
use super::{Reg, num};
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

const DOMAIN: ErrorKind = ErrorKind::Domain;

#[derive(Clone, Copy, PartialEq)]
enum Fold {
    Sum,
    Count,
    CountA,
    Average,
    Get,
    Max,
    Min,
    Product,
    Stdev(bool),
    Var(bool),
}

fn header_text(e: &mut Engine, v: &Value) -> Option<String> {
    match v {
        Value::Text(t) => Some(t.to_string()),
        Value::Number(_) | Value::Bool(_) => e.text(v).ok(),
        _ => None,
    }
}

/// The column of the table that the field argument means.
fn field_column(e: &mut Engine, field: &Value, heads: &[Value], fold: Fold) -> Res<usize> {
    let field = e.scalar(field.clone());
    match field {
        Value::Error(k) => Err(k),
        Value::Text(t) => {
            let want = t.to_lowercase();
            for (i, h) in heads.iter().enumerate() {
                if header_text(e, h).is_some_and(|x| x.to_lowercase() == want) {
                    return Ok(i);
                }
            }
            Err(DOMAIN)
        }
        v @ (Value::Number(_) | Value::Bool(_)) => {
            let n = e.num(&v)?.trunc();
            if n == 0.0 && matches!(fold, Fold::Count | Fold::CountA) {
                // Counting the records: LibreOffice answers it with rules of its own.
                crate::notes::note(crate::notes::DB_FIELD_ZERO);
                return Err(ErrorKind::Unsupported);
            }
            if n >= 1.0 && n <= heads.len() as f64 {
                Ok(n as usize - 1)
            } else {
                Err(ErrorKind::Value)
            }
        }
        _ => Err(DOMAIN),
    }
}

fn database(e: &mut Engine, a: &[Value], fold: Fold) -> Res<Value> {
    for v in [&a[0], &a[2]] {
        if let Value::Error(k) = v {
            return Err(*k);
        }
    }
    let (Value::Ref(_), Value::Ref(_)) = (&a[0], &a[2]) else {
        return Err(DOMAIN);
    };
    let table = e.array_of(a[0].clone());
    let block = e.array_of(a[2].clone());
    let (ncols, nrows) = (table.cols as usize, table.rows as usize);
    let heads: Vec<Value> = (0..ncols).map(|c| table.get(0, c as u32).clone()).collect();
    let col = field_column(e, &a[1], &heads, fold)?;
    if block.rows < 2 {
        return Err(DOMAIN);
    }
    // For every criteria column: the column of the table it tests.
    let mut targets = Vec::new();
    for c in 0..block.cols {
        let h = block.get(0, c).clone();
        let Some(name) = header_text(e, &h) else {
            targets.push(None);
            continue;
        };
        let want = name.to_lowercase();
        let found = heads
            .iter()
            .position(|x| header_text(e, x).is_some_and(|t| t.to_lowercase() == want));
        match found {
            Some(i) => targets.push(Some(i)),
            None => return Err(DOMAIN),
        }
    }
    // The alternatives: per row, the tests of its filled cells.
    let mut alternatives: Vec<Vec<(usize, Criterion)>> = Vec::new();
    for r in 1..block.rows {
        let mut tests = Vec::new();
        for (c, target) in targets.iter().enumerate() {
            let cell = block.get(r, c as u32).clone();
            if let (Some(t), false) = (target, matches!(cell, Value::Empty)) {
                tests.push((*t, Criterion::new(e, &cell)?));
            }
        }
        alternatives.push(tests);
    }
    let coll = e.collator(false);
    let mut picked: Vec<Value> = Vec::new();
    for r in 1..nrows {
        let hit = alternatives.iter().any(|tests| {
            tests
                .iter()
                .all(|(t, crit)| crit.matches(table.get(r as u32, *t as u32), coll))
        });
        if hit {
            picked.push(table.get(r as u32, col as u32).clone());
        }
    }
    let numbers: Vec<f64> = picked
        .iter()
        .filter_map(|v| match v {
            Value::Number(n) => Some(*n),
            _ => None,
        })
        .collect();
    let n = numbers.len() as f64;
    match fold {
        Fold::Sum => num(numbers.iter().sum()),
        Fold::Count => num(n),
        Fold::CountA => num(picked.iter().filter(|v| !matches!(v, Value::Empty)).count() as f64),
        Fold::Average if numbers.is_empty() => Err(ErrorKind::DivZero),
        Fold::Average => num(numbers.iter().sum::<f64>() / n),
        Fold::Get => match picked.as_slice() {
            [Value::Empty] => Err(ErrorKind::Value),
            [one] => Ok(one.clone()),
            [] => Err(ErrorKind::Value),
            _ => Err(DOMAIN),
        },
        Fold::Max => num(numbers
            .iter()
            .copied()
            .fold(None, |m: Option<f64>, x| Some(m.map_or(x, |m| m.max(x))))
            .unwrap_or(0.0)),
        Fold::Min => num(numbers
            .iter()
            .copied()
            .fold(None, |m: Option<f64>, x| Some(m.map_or(x, |m| m.min(x))))
            .unwrap_or(0.0)),
        Fold::Product if numbers.is_empty() => num(0.0),
        Fold::Product => num(numbers.iter().product()),
        Fold::Stdev(sample) | Fold::Var(sample) => {
            // A sample of none is 0, of one is `#NUM!`; a population of none is `#NUM!`.
            if sample && n == 0.0 {
                return num(0.0);
            }
            if n < if sample { 2.0 } else { 1.0 } {
                return Err(ErrorKind::Num);
            }
            let mean = numbers.iter().sum::<f64>() / n;
            let var = numbers.iter().map(|x| (x - mean).powi(2)).sum::<f64>()
                / if sample { n - 1.0 } else { n };
            num(if matches!(fold, Fold::Stdev(_)) {
                var.sqrt()
            } else {
                var
            })
        }
    }
}

pub(super) fn register(r: &mut Reg) {
    macro_rules! database_function {
        ($name:literal, $fold:expr) => {
            r.eager($name, 3, 3, |e, a| database(e, a, $fold))
                .array_args();
        };
    }
    database_function!("DSUM", Fold::Sum);
    database_function!("DCOUNT", Fold::Count);
    database_function!("DCOUNTA", Fold::CountA);
    database_function!("DAVERAGE", Fold::Average);
    database_function!("DGET", Fold::Get);
    database_function!("DMAX", Fold::Max);
    database_function!("DMIN", Fold::Min);
    database_function!("DPRODUCT", Fold::Product);
    database_function!("DSTDEV", Fold::Stdev(true));
    database_function!("DSTDEVP", Fold::Stdev(false));
    database_function!("DVAR", Fold::Var(true));
    database_function!("DVARP", Fold::Var(false));
}
