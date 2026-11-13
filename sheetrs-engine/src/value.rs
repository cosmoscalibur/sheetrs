//! Runtime values shared by the IR, the optimizer and the engine.

use crate::addr::Rect;
use crate::collate::Collator;
use serde::{Deserialize, Serialize};
use std::cmp::Ordering;
use std::sync::Arc;

/// Spreadsheet error values. `Circ` is an engine extension for unresolved
/// circular references (no counterpart in the file formats).
#[derive(
    Debug,
    Clone,
    Copy,
    PartialEq,
    Eq,
    Hash,
    Serialize,
    Deserialize,
    rkyv::Archive,
    rkyv::Serialize,
    rkyv::Deserialize,
)]
pub enum ErrorKind {
    Null,
    DivZero,
    Value,
    Ref,
    Name,
    Num,
    Na,
    Circ,
    /// An engine extension: `LibreOffice` gives a result here that the engine does not reproduce
    /// (see `notes`). The build stops when it meets the case; this is what a run gives when the
    /// case only shows up with the inputs.
    Unsupported,
    /// Internal: an argument outside the domain of the function. It never reaches
    /// a cell: the engine turns it into `#NUM!` or `Libre(502)`, as `LibreOffice` does.
    Domain,
    /// An error only `LibreOffice` has, with its code: `Err:502` (invalid argument), `Err:523` (the
    /// iteration does not converge)… A file format that has no such codes holds it as `#VALUE!`
    /// (an XLSX does, as `LibreOffice`'s own export writes it); an ODS holds the code. It behaves
    /// as `#VALUE!` in every other respect (propagation); `ERROR.TYPE` of it is `#N/A`.
    Libre(u16),
}

impl ErrorKind {
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Null => "#NULL!",
            Self::DivZero => "#DIV/0!",
            Self::Value | Self::Libre(_) => "#VALUE!",
            Self::Ref => "#REF!",
            Self::Name => "#NAME?",
            Self::Num | Self::Domain => "#NUM!",
            Self::Na => "#N/A",
            Self::Circ => "#CIRC!",
            Self::Unsupported => "#NOT_IMPLEMENTED!",
        }
    }

    /// The code `LibreOffice` shows for the error when it is not its Excel name: `Err:502`,
    /// `Err:522` for a circular reference, …
    #[must_use]
    pub fn libre_code(self) -> Option<String> {
        match self {
            Self::Libre(code) => Some(format!("Err:{code}")),
            Self::Circ => Some("Err:522".to_string()),
            _ => None,
        }
    }

    #[must_use]
    pub fn parse(s: &str) -> Option<Self> {
        Some(match s.to_ascii_uppercase().as_str() {
            "#NULL!" => Self::Null,
            "#DIV/0!" => Self::DivZero,
            "#VALUE!" => Self::Value,
            "#REF!" => Self::Ref,
            "#NAME?" => Self::Name,
            "#NUM!" => Self::Num,
            "#N/A" => Self::Na,
            "#CIRC!" | "ERR:522" => Self::Circ,
            "#NOT_IMPLEMENTED!" => Self::Unsupported,
            other => Self::Libre(other.strip_prefix("ERR:")?.parse().ok()?),
        })
    }

    /// Code returned by the `ERROR.TYPE` function (1-based, `#N/A` is 7).
    #[must_use]
    pub fn code(self) -> f64 {
        match self {
            Self::Null => 1.0,
            Self::DivZero => 2.0,
            Self::Value | Self::Libre(_) => 3.0, // `ERROR.TYPE` gives `#N/A` for a `Libre` (see `logic.rs`)
            Self::Ref => 4.0,
            Self::Name => 5.0,
            Self::Num | Self::Domain => 6.0,
            Self::Na => 7.0,
            Self::Circ => 8.0,
            Self::Unsupported => 9.0,
        }
    }
}

/// Rows that follow the stored ones of an [`Array`] and all hold `value`: the part of a range past
/// the data of its sheet, as an operator over it sees it (`Data!A1:A40000*1` has 40,000 rows, the
/// empty ones 0). They are not stored: a function that does not know them gets the array with
/// them stored (see `Engine::array_of`).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Ghost {
    pub rows: u32,
    pub value: Value,
}

/// Rectangular block of scalar values (row-major). Cells are never `Array`
/// or `Ref`.
#[derive(
    Debug,
    Clone,
    PartialEq,
    Serialize,
    Deserialize,
    rkyv::Archive,
    rkyv::Serialize,
    rkyv::Deserialize,
)]
#[rkyv(
    serialize_bounds(
        __S: rkyv::ser::Writer + rkyv::ser::Allocator + rkyv::ser::Sharing,
        <__S as rkyv::rancor::Fallible>::Error: rkyv::rancor::Source,
    ),
    deserialize_bounds(
        __D: rkyv::de::Pooling,
        <__D as rkyv::rancor::Fallible>::Error: rkyv::rancor::Source,
    ),
    bytecheck(bounds(
        __C: rkyv::validation::ArchiveContext + rkyv::validation::SharedContext,
        <__C as rkyv::rancor::Fallible>::Error: rkyv::rancor::Source,
    ))
)]
pub struct Array {
    pub rows: u32,
    pub cols: u32,
    #[rkyv(omit_bounds)]
    pub data: Vec<Value>,
    /// More rows, all equal, after `rows` (none for most arrays). Only a run has them: a program
    /// never holds one, so it is not written.
    #[serde(skip)]
    #[rkyv(with = rkyv::with::Skip)]
    pub ghost: Option<Box<Ghost>>,
    /// The array is what `IF` or `CHOOSE` chose element by element, and the last branch that set a
    /// format in `LibreOffice` was a boolean (`TRUE()`, `NOT(…)`, a comparison…): the cell that
    /// holds its first element shows it as a boolean, whatever the number is. Only the result of a
    /// formula looks at it (`ISLOGICAL`, `INDEX`, `TYPE`… see the elements as numbers), so it is
    /// not part of a program.
    #[serde(skip)]
    #[rkyv(with = rkyv::with::Skip)]
    pub logical: bool,
    /// The array is what `IF` or `CHOOSE` chose element by element (the matrix `LibreOffice` jumps
    /// through). Where a function takes one value, such an array is taken element by element only in
    /// array evaluation; in a formula that is not an array formula the function gets its first
    /// element (`SUM(ABS(IF({1;0},2,3)))` is 2, `SUM(ABS({-1;-2}))` is 3). Only the result of a
    /// formula looks at it, so it is not part of a program.
    #[serde(skip)]
    #[rkyv(with = rkyv::with::Skip)]
    pub jump: bool,
    /// The element that a missing `else` gave shows as `FALSE` in the cell that holds the array: the
    /// array is what `IF`, `CHOOSE`, `IFERROR` or `IFNA` gave itself (any other array shows an empty
    /// text there, and `TRANSPOSE` a 0). Only the result of a formula looks at it.
    #[serde(skip)]
    #[rkyv(with = rkyv::with::Skip)]
    pub jump_false: bool,
}

impl Array {
    #[must_use]
    pub fn new(rows: u32, cols: u32, data: Vec<Value>) -> Self {
        debug_assert_eq!(data.len(), rows as usize * cols as usize);
        Self {
            rows,
            cols,
            data,
            ghost: None,
            logical: false,
            jump: false,
            jump_false: false,
        }
    }

    /// The array with `rows` more rows of `value` after its own, not stored.
    #[must_use]
    pub fn with_ghost(mut self, rows: u32, value: Value) -> Self {
        self.ghost = (rows > 0).then(|| Box::new(Ghost { rows, value }));
        self
    }

    /// The array marked as the result of branches that left a boolean format.
    #[must_use]
    pub fn with_logical(mut self, logical: bool) -> Self {
        self.logical = logical;
        self
    }

    /// The array marked as what `IF` or `CHOOSE` chose element by element.
    #[must_use]
    pub fn with_jump(mut self, jump: bool) -> Self {
        self.jump = jump;
        self
    }

    /// The array marked as what `IF`, `CHOOSE`, `IFERROR` or `IFNA` gave: jump (see
    /// [`with_jump`](Self::with_jump)) and showing `FALSE` for the element a missing `else` gave.
    #[must_use]
    pub fn with_jump_false(mut self) -> Self {
        self.jump = true;
        self.jump_false = true;
        self
    }

    /// The rows the array has, the ones not stored included.
    #[must_use]
    pub fn logical_rows(&self) -> u64 {
        u64::from(self.rows) + self.ghost.as_ref().map_or(0, |g| u64::from(g.rows))
    }

    /// The value at a row of the array, the rows that are not stored included (`r` must be a row
    /// of the array).
    #[must_use]
    pub fn logical(&self, r: u32, c: u32) -> &Value {
        match &self.ghost {
            Some(g) if r >= self.rows => &g.value,
            _ => self.get(r, c),
        }
    }

    /// The array with its rows after the stored ones stored too.
    #[must_use]
    pub fn expanded(&self) -> Self {
        let Some(g) = &self.ghost else {
            return Self::new(self.rows, self.cols, self.data.clone());
        };
        let rows = self.rows + g.rows;
        let mut data = Vec::with_capacity(rows as usize * self.cols as usize);
        data.extend(self.data.iter().cloned());
        data.resize(rows as usize * self.cols as usize, g.value.clone());
        Self::new(rows, self.cols, data)
    }

    #[must_use]
    pub fn filled(rows: u32, cols: u32, v: Value) -> Self {
        Self::new(rows, cols, vec![v; rows as usize * cols as usize])
    }

    #[must_use]
    pub fn get(&self, r: u32, c: u32) -> &Value {
        &self.data[r as usize * self.cols as usize + c as usize]
    }
}

#[derive(
    Debug,
    Clone,
    PartialEq,
    Default,
    Serialize,
    Deserialize,
    rkyv::Archive,
    rkyv::Serialize,
    rkyv::Deserialize,
)]
#[rkyv(
    serialize_bounds(
        __S: rkyv::ser::Writer + rkyv::ser::Allocator + rkyv::ser::Sharing,
        <__S as rkyv::rancor::Fallible>::Error: rkyv::rancor::Source,
    ),
    deserialize_bounds(
        __D: rkyv::de::Pooling,
        <__D as rkyv::rancor::Fallible>::Error: rkyv::rancor::Source,
    ),
    bytecheck(bounds(
        __C: rkyv::validation::ArchiveContext + rkyv::validation::SharedContext,
        <__C as rkyv::rancor::Fallible>::Error: rkyv::rancor::Source,
    ))
)]
pub enum Value {
    #[default]
    Empty,
    Number(f64),
    Bool(bool),
    Text(Arc<str>),
    Error(ErrorKind),
    Array(#[rkyv(omit_bounds)] Arc<Array>),
    /// A reference. Only produced while evaluating; never stored in a cell.
    Ref(Rect),
}

/// The empty text that a branch `IF` does not have gives for an element of an array (`IF({0;1},1)`:
/// the element of the first row): in `LibreOffice` an empty text that is not counted by `COUNTA`, and
/// that a cell holding the array shows as `FALSE`. It is the one string every such element shares, so
/// that it can be told from the empty text of a formula (`Arc::ptr_eq`): nothing else sees a
/// difference (`LEN` is 0, `TYPE` 2, `=""` true, `+0` is `#VALUE!`).
static EMPTY_PATH: std::sync::LazyLock<Arc<str>> = std::sync::LazyLock::new(|| Arc::from(""));

/// The same element as an operator or a function applied to each element of an array reads it
/// (`LEN(IF({0;1},1))`): `LibreOffice` reads it as the text `FALSE` where a text is taken (`LEN` is
/// 5, `&` gives `FALSE`), as the number 0 where a number is (`+`, `ABS`), as empty in a comparison
/// (`=""` and `=0` are both true) and as a text for `ISTEXT`. One string shared by every such
/// element, so that the places that read it otherwise can tell it (`Arc::ptr_eq`).
static ELEMENT_PATH: std::sync::LazyLock<Arc<str>> =
    std::sync::LazyLock::new(|| Arc::from("FALSE"));

impl Value {
    pub fn text(s: impl AsRef<str>) -> Self {
        Self::Text(Arc::from(s.as_ref()))
    }

    /// The element of an array that the missing branch of an `IF` gives (see `EMPTY_PATH`).
    #[must_use]
    pub fn empty_path() -> Self {
        Self::Text(Arc::clone(&EMPTY_PATH))
    }

    /// The element the missing branch of an `IF` gave, as a call applied to each element reads it
    /// (see `ELEMENT_PATH`).
    #[must_use]
    pub fn element_path() -> Self {
        Self::Text(Arc::clone(&ELEMENT_PATH))
    }

    /// Whether this is [`element_path`](Self::element_path).
    #[must_use]
    pub fn is_element_path(&self) -> bool {
        matches!(self, Self::Text(t) if Arc::ptr_eq(t, &ELEMENT_PATH))
    }

    /// Whether this is the element the missing branch of an `IF` gives, not any other empty text.
    #[must_use]
    pub fn is_empty_path(&self) -> bool {
        matches!(self, Self::Text(t) if Arc::ptr_eq(t, &EMPTY_PATH))
    }

    #[must_use]
    pub fn is_error(&self) -> bool {
        matches!(self, Self::Error(_))
    }

    #[must_use]
    pub fn as_error(&self) -> Option<ErrorKind> {
        match self {
            Self::Error(e) => Some(*e),
            _ => None,
        }
    }

    /// Rank used by the cross-type ordering: numbers (booleans are the numbers 1 and 0 in
    /// `LibreOffice`) < text.
    fn type_rank(&self) -> u8 {
        match self {
            Self::Empty | Self::Number(_) | Self::Bool(_) => 0,
            Self::Text(_) => 1,
            _ => 3,
        }
    }

    /// The number a value is for ordering and matching: a boolean is 1 or 0, an empty cell 0.
    #[must_use]
    pub fn as_ordinal(&self) -> Option<f64> {
        match self {
            Self::Number(n) => Some(*n),
            Self::Bool(b) => Some(f64::from(u8::from(*b))),
            Self::Empty => Some(0.0),
            _ => None,
        }
    }
}

/// Text collators for a locale: one that ignores case (accents still count)
/// and one that breaks ties by case, lowercase first. See `crate::collate`.
pub type Coll = Collator;

#[derive(Debug)]
pub struct Collators {
    ci: Coll,
    cs: Coll,
}

impl Collators {
    /// The collators of the language of `tag`.
    #[must_use]
    pub fn new(tag: &str) -> Self {
        Self {
            ci: Collator::new(tag, false),
            cs: Collator::new(tag, true),
        }
    }

    /// `case_sensitive`: case decides ties instead of being ignored.
    #[must_use]
    pub fn get(&self, case_sensitive: bool) -> &Coll {
        if case_sensitive { &self.cs } else { &self.ci }
    }
}

/// Spreadsheet ordering of two scalars. `Empty` takes the type of the other
/// operand.
#[must_use]
pub fn compare(a: &Value, b: &Value, col: &Coll) -> Ordering {
    let (a, b) = match (a, b) {
        (Value::Empty, Value::Text(_)) => (&Value::Text(Arc::from("")), b),
        (Value::Text(_), Value::Empty) => (a, &Value::Text(Arc::from(""))),
        _ => (a, b),
    };
    match (a, b) {
        (Value::Text(x), Value::Text(y)) => col.compare(x, y),
        (x, y) if x.as_ordinal().is_some() && y.as_ordinal().is_some() => {
            let (x, y) = (x.as_ordinal().unwrap_or(0.0), y.as_ordinal().unwrap_or(0.0));
            x.partial_cmp(&y).unwrap_or(Ordering::Equal)
        }
        _ => a.type_rank().cmp(&b.type_rank()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn collation_orders_accents_and_case_as_the_locale_does() {
        let en = Collators::new("en");
        let es = Collators::new("es-ES");
        let (ci, cs) = (en.get(false), en.get(true));
        assert_eq!(
            ci.compare("á", "b"),
            Ordering::Less,
            "accents do not jump past letters"
        );
        assert_eq!(ci.compare("a", "á"), Ordering::Less);
        assert_eq!(ci.compare("ABC", "abc"), Ordering::Equal);
        assert_eq!(cs.compare("abc", "ABD"), Ordering::Less);
        assert_eq!(cs.compare("a", "A"), Ordering::Less);
        assert_eq!(ci.compare("Ñandú", "ñANDÚ"), Ordering::Equal);
        // English: ñ is n with a tilde. Spanish: a letter between n and o.
        assert_eq!(ci.compare("ña", "nb"), Ordering::Less);
        assert_eq!(es.get(false).compare("ña", "nz"), Ordering::Greater);
        assert_eq!(es.get(false).compare("ña", "oa"), Ordering::Less);
        assert_eq!(es.get(false).compare("nz", "ñ"), Ordering::Less);
        assert_eq!(es.get(false).compare("ñ", "Ñ"), Ordering::Equal);
    }
}
