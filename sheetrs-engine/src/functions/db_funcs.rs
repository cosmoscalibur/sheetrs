//! The database functions `DSUM`, `DCOUNT`, `DCOUNTA`, `DAVERAGE`, `DGET`, `DMAX`, `DMIN`,
//! `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DVAR` and `DVARP`: a table whose first row names its fields, a
//! field (a name or a position) and a block of criteria (a row of names and one row per
//! alternative, whose cells all have to hold). What `LibreOffice` 26.8 answers for the arguments
//! outside that shape is what `tests/assets/semantics.*` holds.

use super::stats::Criterion;
use super::{Reg, num};
use crate::engine::{Engine, Res};
use crate::ir::Expr;
use crate::options::SearchMode;
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
fn field_column(
    e: &mut Engine,
    field: &Value,
    field_is_range: bool,
    heads: &[Value],
    fold: Fold,
    table_literal: bool,
) -> Res<usize> {
    // A field written as a range, even of one cell, is outside the domain.
    if field_is_range {
        return Err(DOMAIN);
    }
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
            } else if n < 1.0 && table_literal {
                // A position before the first column of a table written in the formula is not in
                // the domain; in a range it is a `#VALUE!`, as one after the last column is.
                Err(DOMAIN)
            } else {
                Err(ErrorKind::Value)
            }
        }
        _ => Err(DOMAIN),
    }
}

/// How the arguments were written: a reference or an array written in the formula for the table
/// and for the criteria, a cell or a range for the criteria and the field.
// One flag per argument and per way it can be written: the arguments are told apart by the shape.
#[allow(clippy::struct_excessive_bools)]
#[derive(Clone, Copy)]
struct Shape {
    /// The table is an array written in the formula (`{"Name","Age";"Ann",30}`).
    table_literal: bool,
    /// The criteria are an array written in the formula.
    criteria_literal: bool,
    /// The criteria are a single cell written as such (`I1`), which is no block.
    criteria_cell: bool,
    /// The field is written as a range.
    field_is_range: bool,
}

fn database(e: &mut Engine, a: &[Value], fold: Fold, shape: Shape) -> Res<Value> {
    for v in [&a[0], &a[2]] {
        if let Value::Error(k) = v {
            return Err(*k);
        }
    }
    // A table or criteria is a range, or an array that is written in the formula (what a formula
    // computes is no table there).
    let given = |v: &Value, literal: bool| match v {
        Value::Ref(_) => true,
        Value::Array(_) => literal,
        _ => false,
    };
    if !given(&a[0], shape.table_literal) || !given(&a[2], shape.criteria_literal) {
        return Err(DOMAIN);
    }
    let table = e.array_of(a[0].clone());
    let block = e.array_of(a[2].clone());
    // The reference may reach past what the sheet holds: the names it takes there are blank. A
    // single cell is no block (`I1`); a row of names alone (`I1:J1`, or `I1:I1`) is one with no
    // condition.
    if let Value::Ref(block_ref) = &a[2]
        && block_ref.cols() > block.cols
    {
        return Err(DOMAIN);
    }
    if shape.criteria_cell {
        return Err(DOMAIN);
    }
    let (criteria_literal, field_is_range) = (shape.criteria_literal, shape.field_is_range);
    let (ncols, nrows) = (table.cols as usize, table.rows as usize);
    let heads: Vec<Value> = (0..ncols).map(|c| table.get(0, c as u32).clone()).collect();
    let col = field_column(e, &a[1], field_is_range, &heads, fold, shape.table_literal)?;
    // For every criteria column: the column of the table it tests.
    let mut targets = Vec::new();
    for c in 0..block.cols {
        let h = block.get(0, c).clone();
        // A criteria column must be headed by a name of the table, as `LibreOffice` asks.
        let Some(name) = header_text(e, &h) else {
            return Err(DOMAIN);
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
    // The text of a criterion is matched as the document says, except in a table written in the
    // formula: there it is the whole text of a cell, with no wildcard or expression.
    let mut search = e.opts.search;
    if shape.table_literal {
        search.mode = SearchMode::Literal;
        search.whole_cell = true;
    }
    // The alternatives: per row, the tests of its filled cells.
    let mut alternatives: Vec<Vec<(usize, Criterion)>> = Vec::new();
    for r in 1..block.rows {
        let mut tests = Vec::new();
        for (c, target) in targets.iter().enumerate() {
            let cell = block.get(r, c as u32).clone();
            // In an array written in the formula a number, a boolean or an empty text is no
            // condition, as a blank cell is none.
            let blank = match &cell {
                Value::Empty => true,
                Value::Number(_) | Value::Bool(_) => criteria_literal,
                Value::Text(t) => criteria_literal && t.is_empty(),
                _ => false,
            };
            if let (Some(t), false) = (target, blank) {
                tests.push((*t, Criterion::with_search(e, &cell, search)?));
            }
        }
        alternatives.push(tests);
    }
    // A row with no condition is no alternative; a block with no condition at all (the names
    // alone, or only blank rows) takes every record.
    // (In a table written in the formula no record is taken then.)
    alternatives.retain(|tests| !tests.is_empty());
    if alternatives.is_empty() && !shape.table_literal {
        alternatives.push(Vec::new());
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
            [Value::Empty] | [] => Err(ErrorKind::Value),
            [one] => Ok(one.clone()),
            _ => Err(ErrorKind::Libre(502)),
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
            // Lazy, to tell a cell from a range of one cell in the field and the criteria.
            r.lazy($name, 3, 3, |e, a| {
                let literal = |x: &Expr| matches!(x, Expr::Const(Value::Array(_)));
                let shape = Shape {
                    table_literal: literal(&a[0]),
                    criteria_literal: literal(&a[2]),
                    criteria_cell: matches!(a[2], Expr::Cell(_)),
                    field_is_range: matches!(a[1], Expr::Range(_)),
                };
                // As the arguments of any function: an expression over a range is met by the cell
                // of the formula unless the formula is an array formula.
                let values: Vec<Value> = a.iter().map(|x| e.eval(x)).collect();
                database(e, &values, $fold, shape)
            });
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
