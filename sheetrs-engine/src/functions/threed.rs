//! A reference over several sheets (`Data!A1:Wide!A3`, or `Data:Wide!A1:A3` as an XLSX writes it) is the
//! same area on each sheet from the first to the last, the sheets in between included, in the order
//! of the workbook. What a function does with it is for `LibreOffice` to say, and it says three things
//! (`scripts/probe_3d.py` asks it of every function and writes [`THREED`]): the function goes through
//! the cells of every sheet (`SUM`, `COUNT`, `MAX`, `LARGE`…), or it refuses it with an error of its own
//! (`Err:504` for `SUMIF`, `VLOOKUP`, `MATCH`, `SUMPRODUCT`…, `#REF!` for `INDEX`), or it takes a value
//! there and a reference over several sheets is not one (`ABS`, an operator: `#VALUE!`).
//!
//! The compiler lowers such a reference given to a function by asking [`outcome`]: a function that
//! goes through the cells gets the area of each sheet as an argument of its own (a function with any
//! number of arguments) or one array of all of them (`_3DA`), and the call becomes the error when
//! the function refuses it. A reference over several sheets anywhere else is `#VALUE!`.

use super::threed_table::THREED;
use super::{Reg, VAR};
use crate::engine::Res;
use crate::ir::Ns;
use crate::value::{Array, ErrorKind, Value};

/// What a function does with a reference over several sheets in one of its arguments.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Outcome {
    /// It goes through the cells of every sheet.
    Iterate,
    /// It refuses it: the call is this error.
    Error(ErrorKind),
}

/// Whether a function takes any number of arguments, and so the area of each sheet can be one of them.
#[must_use]
pub fn is_variadic(def: &super::FuncDef) -> bool {
    def.max == VAR
}

/// What the function `name` of the namespace `ns` does with a reference over several sheets as its
/// argument number `arg` (from 0).
#[must_use]
pub fn outcome(name: &str, ns: Ns, arg: usize) -> Outcome {
    let Ok(i) = THREED.binary_search_by(|&(n, s, _)| (n, s).cmp(&(name, ns.as_str()))) else {
        return Outcome::Iterate;
    };
    let codes = THREED[i].2.as_bytes();
    match codes.get(arg).or_else(|| codes.last()) {
        Some(b'4') => Outcome::Error(ErrorKind::Libre(504)),
        Some(b'2') => Outcome::Error(ErrorKind::Libre(502)),
        Some(b'v') => Outcome::Error(ErrorKind::Value),
        Some(b'r') => Outcome::Error(ErrorKind::Ref),
        Some(b'n') => Outcome::Error(ErrorKind::Name),
        Some(b'a') => Outcome::Error(ErrorKind::Na),
        _ => Outcome::Iterate,
    }
}

pub(super) fn register(r: &mut Reg) {
    // The areas of the sheets as one column of values (what a function that takes one array sees).
    r.eager("_3DA", 1, VAR, |e, a| -> Res<Value> {
        let mut data = Vec::new();
        for v in a {
            match v {
                Value::Ref(rect) if rect.is_single() => data.push(e.cell_value(rect.top_left())),
                Value::Ref(rect) => data.extend(e.range_array(*rect).data.iter().cloned()),
                other => data.push(other.clone()),
            }
        }
        let rows = data.len() as u32;
        Ok(Value::Array(std::sync::Arc::new(Array::new(rows, 1, data))))
    });
}
