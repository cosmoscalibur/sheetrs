//! The number format a formula cell takes from its result.
//!
//! `LibreOffice` gives a formula cell whose format is `General` the format its result has: `=A1`
//! over a date shows a date, `=A1+1` the default date format, `=A5*2` over a currency a currency.
//! This module reproduces the rules of its interpreter (observed in 26.8 with probes of every
//! operator and function over every family of format; the probes are the fixture
//! `sheetrs-io/tests/assets/inherit.*` and the oracle made from it):
//!
//! **Which cells.** Only a formula cell whose own format is `General` (a style the file gives wins),
//! and only when `LibreOffice` interprets it while it loads the file:
//!
//! - in an **ODS**, every formula cell without a format (the import asks for the format of the result);
//! - in an **XLSX**, the cells that are recalculated when the file is opened: those that call a
//!   volatile function (`RAND`, `RANDBETWEEN`, `RANDARRAY`, `NOW`, `TODAY`, `INDIRECT`, `OFFSET`,
//!   `CELL`) and, through their references, those that depend on one. The others keep what the file
//!   says, `General`.
//!
//! **Rules.** A result is a *reference* or a *format*. The format of a reference is the one of the
//! cell (its own, or the one it was given by these rules).
//!
//! | Expression | Format |
//! | ---------- | ------ |
//! | a cell, `INDIRECT`, `OFFSET`, `INDEX`, `CHOOSE`, `IF` (the branch taken), `IFERROR`, `IFNA` | that of the cell the reference gives |
//! | `+` and `-` | by the classes of the two operands, see `Inherit::add` |
//! | `*` | the currency of an operand, else `General` |
//! | `/` | the currency of the first operand when the second is not one, else `General` |
//! | `^`, `&`, `-x`, `x^y`, functions not listed | `General` |
//! | `x%` | the currency of `x`, else `0.00%` |
//! | `SUM` | the format of the last cell of the first argument that holds a number |
//! | `MIN`, `MAX`, `AVERAGE`, `PRODUCT`, `SUMSQ`, `AVERAGEA`, `MAXA`, `MINA`, `SUBTOTAL`, `AGGREGATE` | the format of the first cell of the first argument that holds a number |
//! | comparisons, `TRUE`, `FALSE`, `NOT`, `AND`, `OR`, `XOR`, `IS…`, `EXACT` | `BOOLEAN` |
//! | `DATE`, `TODAY` / `NOW` / `TIME` | the default date format / date and time / time of day |
//!
//! The aggregates keep the format of a cell with a date, a time, a currency, a scientific number or
//! a fraction, turn a percentage into the default percentage and a number format (`#,##0.00`, `0.000`)
//! into `General`; `MIN` and `MAX` also keep `BOOLEAN`. The first and the last cell are those that
//! hold a number, in the order of the columns; an argument with none (text, empty cells) is passed over
//! and the next one decides, and a number written in the formula (`SUM(1,A1)`) is an argument with
//! the format `General`. The text format (`@`) is never taken from a cell: a reference to a cell
//! formatted as text gives `General`, as does an aggregate of one. Every other kind of cell and result
//! is `General`.
//!
//! **What chooses a branch or a cell.** An `IF`, a `CHOOSE` or an `INDIRECT` takes the format of the
//! branch or the cell its condition, index or text names, which is worked out here from constants,
//! cells (their value as the file caches it), defined names and these functions: `IF`, `NOT`, `AND`,
//! `OR`, `TRUE`, `FALSE`, `ISNUMBER`, `ISTEXT`, `ISBLANK`, `LEN`, `ABS`, `INT`, `ROUND`, `MOD`, `ROW`,
//! `COLUMN`, `CONCATENATE`, `CONCAT`, `UPPER`, `LOWER`, `LEFT`, `RIGHT`, `ADDRESS` (two arguments),
//! `SUM`, `COUNT`, `MAX`, `MIN`, and the operators. A **defined name** stands for what its definition
//! says (a cell, a range, a constant) wherever a formula uses it, and `INDIRECT` takes a text that is
//! the name of a range; a name defined in more than one scope is not resolved (which one a formula
//! means depends on its sheet, and the file does not say). The probes are the fixture
//! `sheetrs-io/tests/assets/inherit_more.*` and the oracle made from it.
//!
//! **Limits** (reported, not approximated): an `IF`, `CHOOSE` or `INDIRECT` whose condition, index or
//! text depends on anything else (a function that is not in the list above, a volatile function, a
//! cell without a cached value) gives `General`, as does a name that is not a reference or a
//! constant (a formula, a table column); an ODS cell whose result is a time and has no format shows
//! `HH:MM:SS AM/PM` in `LibreOffice`, which needs the type of the value, and it is not read; only the
//! default formats of `en-US` are written (the date order of the locale is followed).

use crate::text::TextContext;
use sheetrs_formula::addr::{parse_area, same_name, split_sheet};
use sheetrs_formula::ast::{BinOp, Call, Expr, RefArea, Reference, UnOp};
use sheetrs_io::Book;
use sheetrs_io::reader::{
    Cell, CellSink, CellValue, Format, Properties, Selection, SheetInfo, Window,
};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::rc::Rc;

/// The format code of a result; `None` is `General`.
type Code = Option<Rc<str>>;

/// A cell of the workbook as the inference needs it.
type Key = (u16, u32, u32);

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Class {
    /// `General`, text, and the formats that no rule treats apart: scientific, fractions.
    Plain,
    /// A number format that is not `General` (`0.00`, `#,##0.00`, `"x"0`).
    Number,
    Percent,
    /// A scientific number (`0.0E+0`): aggregates show it as `0.00E+00`.
    Scientific,
    Currency,
    Date,
    DateTime,
    Time,
    Boolean,
}

fn strip_literals(code: &str) -> String {
    let mut out = String::new();
    let mut chars = code.chars();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for q in chars.by_ref() {
                    if q == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                chars.next();
            }
            c => out.push(c),
        }
    }
    out
}

/// Whether a format code is the text format (`@`, with or without the language mark): a formula
/// never takes it from a cell, its result is `General`.
fn is_text_format(code: &str) -> bool {
    let mut rest = code;
    let mut kept = String::new();
    while let Some(c) = rest.chars().next() {
        if c == '[' {
            match rest.find(']') {
                Some(end) => rest = &rest[end + 1..],
                None => return false,
            }
        } else {
            kept.push(c);
            rest = &rest[c.len_utf8()..];
        }
    }
    kept.contains('@') && kept.chars().all(|c| matches!(c, '@' | ';' | ' '))
}

/// The class of a format code, in the terms of the rules.
fn class_of(code: &Code) -> Class {
    let Some(code) = code else {
        return Class::Plain;
    };
    let lower = code.to_ascii_lowercase();
    if lower == "boolean" || lower == "\"true\";\"true\";\"false\"" {
        return Class::Boolean;
    }
    let plain = strip_literals(code);
    // The currency of a code written `[$€-407]` or `[$USD]` (not the language mark `[$-409]`), or a
    // `$` that is not quoted.
    let mut currency = plain.contains('$') && !plain.replace("[$-", "").is_empty() && {
        let mut rest = plain.as_str();
        let mut found = false;
        while let Some(at) = rest.find('$') {
            let before_dash = rest[at + 1..].starts_with('-');
            let in_bracket = rest[..at].ends_with('[');
            if !(in_bracket && before_dash) {
                found = true;
                break;
            }
            rest = &rest[at + 1..];
        }
        found
    };
    currency |= plain.chars().any(|c| matches!(c, '€' | '£' | '¥'));
    if currency {
        return Class::Currency;
    }
    if plain.contains('%') {
        return Class::Percent;
    }
    let mut letters = String::new();
    let mut elapsed = false;
    let mut chars = plain.chars();
    while let Some(c) = chars.next() {
        if c == '[' {
            let inner: String = chars.by_ref().take_while(|&q| q != ']').collect();
            let inner = inner.to_ascii_lowercase();
            if !inner.is_empty() && inner.chars().all(|q| matches!(q, 'h' | 'm' | 's')) {
                elapsed = true;
                letters.push(inner.chars().next().unwrap_or('h'));
            }
        } else if c.is_ascii_alphabetic() {
            letters.push(c.to_ascii_lowercase());
        }
    }
    let letters = letters
        .replace("general", "")
        .replace("am/pm", "")
        .replace("a/p", "");
    let has = |c: char| letters.contains(c);
    let scientific = lower.contains("e+") || lower.contains("e-");
    let date =
        !scientific && (has('y') || has('d') || has('n') || (has('m') && !(has('h') || has('s'))));
    let time = has('h') || has('s') || elapsed;
    if date && time {
        return Class::DateTime;
    }
    if date {
        return Class::Date;
    }
    if time && !scientific {
        return Class::Time;
    }
    if scientific {
        return Class::Scientific;
    }
    if plain.contains("?/") || plain.contains("#/") || plain.contains('@') {
        return Class::Plain;
    }
    if plain.contains('0') || plain.contains('#') || plain.contains('?') {
        return Class::Number;
    }
    Class::Plain
}

/// A constant worked out from literals and cells, for what chooses a branch or a cell.
#[derive(Clone, Debug, PartialEq)]
enum Const {
    Num(f64),
    Text(String),
    Bool(bool),
}

impl Const {
    fn number(&self) -> Option<f64> {
        match self {
            Self::Num(n) => Some(*n),
            Self::Bool(b) => Some(f64::from(u8::from(*b))),
            Self::Text(_) => None,
        }
    }

    fn truth(&self) -> Option<bool> {
        match self {
            Self::Bool(b) => Some(*b),
            Self::Num(n) => Some(*n != 0.0),
            Self::Text(_) => None,
        }
    }

    fn text(&self) -> String {
        match self {
            Self::Text(t) => t.clone(),
            Self::Num(n) => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    n.to_string()
                }
            }
            Self::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        }
    }
}

/// What an expression gives, for the rules.
#[derive(Clone, Debug)]
enum Val {
    /// A format.
    Fmt(Code),
    /// A reference to cells: `(sheet, row0, col0, row1, col1)`.
    Ref(u16, u32, u32, u32, u32),
}

/// Number formats of the formula cells that take one of their result, found for a read of cells.
pub struct Inference<'a> {
    source: &'a Book,
    infos: &'a [SheetInfo],
    text: &'a TextContext,
    /// The cells that were read, ordered by position, with the number format they carry.
    cells: &'a [(u16, u32, u32, Cell)],
    /// The cells read for the references that lie outside them.
    extra: BTreeMap<Key, Cell>,
    loaded: Vec<(u16, u32, u32, u32, u32)>,
    /// The windows read so far for each sheet, and the sheets read whole.
    windows: HashMap<u16, usize>,
    whole: HashSet<u16>,
    ods: bool,
    formats: HashMap<Key, Code>,
    recalculated: HashMap<Key, bool>,
    busy: HashSet<Key>,
    /// The defined names, by upper-case name, read when a formula first needs one.
    names: Option<HashMap<String, Expr>>,
    /// The names being resolved: a name that refers to itself stands for nothing.
    names_busy: HashSet<String>,
}

/// Collects the cells of a read that the search has not kept (`known` are in order): `extra` and
/// the cells of the search never hold the same position.
struct Gather<'a> {
    out: &'a mut BTreeMap<Key, Cell>,
    known: &'a [(u16, u32, u32, Cell)],
}

impl CellSink for Gather<'_> {
    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell) {
        let key = (sheet, at.0, at.1);
        if (!matches!(cell.value, CellValue::Empty) || cell.formula.is_some())
            && self
                .known
                .binary_search_by_key(&key, |(s, r, c, _)| (*s, *r, *c))
                .is_err()
        {
            self.out.insert(key, cell);
        }
    }
}

/// The windows read for a sheet before it is read whole: every window is a pass over the file, so
/// a workbook whose formulas refer to many different ranges would read it once per range.
const WINDOWS_BEFORE_WHOLE: usize = 8;

const VOLATILE: [&str; 8] = [
    "RAND",
    "RANDBETWEEN",
    "RANDARRAY",
    "NOW",
    "TODAY",
    "INDIRECT",
    "OFFSET",
    "CELL",
];

fn function_name(raw: &str) -> String {
    let upper = raw.to_ascii_uppercase();
    for prefix in ["_XLFN.", "_XLWS.", "COM.MICROSOFT.", "ORG.OPENOFFICE."] {
        if let Some(rest) = upper.strip_prefix(prefix) {
            return rest.to_string();
        }
    }
    upper
}

#[allow(clippy::unnecessary_wraps)]
fn code(s: &str) -> Code {
    Some(Rc::from(s))
}

impl<'a> Inference<'a> {
    #[must_use]
    pub fn new(
        source: &'a Book,
        infos: &'a [SheetInfo],
        cells: &'a [(u16, u32, u32, Cell)],
        text: &'a TextContext,
    ) -> Self {
        Self {
            source,
            infos,
            text,
            cells,
            extra: BTreeMap::new(),
            loaded: Vec::new(),
            windows: HashMap::new(),
            whole: HashSet::new(),
            ods: source.format() == Format::Ods,
            formats: HashMap::new(),
            recalculated: HashMap::new(),
            busy: HashSet::new(),
            names: None,
            names_busy: HashSet::new(),
        }
    }

    /// What a defined name stands for: a reference or a constant, as the formulas write it. A name
    /// the workbook defines in more than one scope is not resolved (which one a formula means
    /// depends on its sheet, and the file does not say).
    fn name_expr(&mut self, name: &str) -> Option<Expr> {
        if self.names.is_none() {
            let mut map = HashMap::new();
            if let Ok((names, conflicts)) = self.source.defined_names() {
                for (name, definition) in names {
                    if conflicts.iter().any(|c| same_name(&c.name, &name)) {
                        continue;
                    }
                    if let Ok(expr) = sheetrs_formula::parse(definition.trim_start_matches('=')) {
                        map.insert(name.to_uppercase(), expr);
                    }
                }
            }
            self.names = Some(map);
        }
        self.names.as_ref()?.get(&name.to_uppercase()).cloned()
    }

    /// Works `f` out on what the name stands for, unless it leads back to itself.
    fn through_name<T>(
        &mut self,
        name: &str,
        f: impl FnOnce(&mut Self, &Expr) -> Option<T>,
    ) -> Option<T> {
        let expr = self.name_expr(name)?;
        let key = name.to_uppercase();
        if !self.names_busy.insert(key.clone()) {
            return None;
        }
        let out = f(self, &expr);
        self.names_busy.remove(&key);
        out
    }

    /// The format code the formula cell at `at` takes from its result, when it takes one.
    pub fn format_of(&mut self, at: Key, cell: &Cell) -> Option<String> {
        let formula = cell.formula.as_deref()?;
        if !own_is_general(cell) {
            return None;
        }
        let code = self.formula_format(at, formula)?;
        code.map(|c| c.to_string())
    }

    fn sheet_name(&self, sheet: u16) -> &str {
        self.infos
            .iter()
            .find(|i| i.index == sheet)
            .map_or("", |i| i.name.as_str())
    }

    fn sheet_index(&self, name: &str) -> Option<u16> {
        self.infos
            .iter()
            .find(|i| same_name(&i.name, name))
            .map(|i| i.index)
    }

    /// The cell at `at`, from what was read or from what was read for the references.
    fn cell(&self, at: Key) -> Option<&Cell> {
        if let Ok(i) = self
            .cells
            .binary_search_by_key(&at, |(s, r, c, _)| (*s, *r, *c))
        {
            return Some(&self.cells[i].3);
        }
        self.extra.get(&at)
    }

    /// Makes sure the cells of a rectangle are known (they are read once, with their formulas).
    fn ensure(&mut self, sheet: u16, r0: u32, c0: u32, r1: u32, c1: u32) {
        if self.whole.contains(&sheet)
            || self
                .loaded
                .iter()
                .any(|&(s, a, b, c, d)| s == sheet && a <= r0 && b <= c0 && c >= r1 && d >= c1)
        {
            return;
        }
        let count = self.windows.entry(sheet).or_insert(0);
        *count += 1;
        let (r0, c0, r1, c1) = if *count > WINDOWS_BEFORE_WHOLE {
            self.whole.insert(sheet);
            (0, 0, u32::MAX, u32::MAX)
        } else {
            (r0, c0, r1, c1)
        };
        let name = self.sheet_name(sheet).to_string();
        let selection = Selection::new(Properties {
            value: true,
            formula: true,
            number_format: true,
            comment: false,
        })
        .window(name, Window::new(r0, c0, r1, c1));
        let mut gathered = BTreeMap::new();
        if self
            .source
            .selection_into(
                &selection,
                &mut Gather {
                    out: &mut gathered,
                    known: self.cells,
                },
            )
            .is_ok()
        {
            self.extra.extend(gathered);
        }
        self.loaded.push((sheet, r0, c0, r1, c1));
    }

    /// The cells with a number in a rectangle, in the order of the columns.
    fn numbers_in(&mut self, sheet: u16, r0: u32, c0: u32, r1: u32, c1: u32) -> Vec<Key> {
        self.ensure(sheet, r0, c0, r1, c1);
        let mut found: Vec<Key> = Vec::new();
        let from = self
            .cells
            .partition_point(|(s, r, _, _)| (*s, *r) < (sheet, r0));
        for (s, r, c, cell) in &self.cells[from..] {
            if *s != sheet || *r > r1 {
                break;
            }
            if (c0..=c1).contains(c) && is_number(cell) {
                found.push((*s, *r, *c));
            }
        }
        for (&(s, r, c), cell) in self.extra.range((sheet, r0, 0)..=(sheet, r1, u32::MAX)) {
            if s == sheet && (c0..=c1).contains(&c) && is_number(cell) {
                found.push((s, r, c));
            }
        }
        found.sort_by_key(|&(s, r, c)| (s, c, r));
        found
    }

    /// The format a cell shows: its own, or the one it takes from its result.
    fn effective(&mut self, at: Key) -> Code {
        if let Some(known) = self.formats.get(&at) {
            return known.clone();
        }
        let cell = self.cell(at).cloned()?;
        let own = own_code(&cell);
        let result = if own.is_some() {
            own
        } else if matches!(cell.value, CellValue::Boolean(_)) {
            code("BOOLEAN")
        } else if let Some(formula) = cell.formula.as_deref() {
            self.formula_format(at, formula).flatten()
        } else {
            None
        };
        self.formats.insert(at, result.clone());
        result
    }

    /// `Some(code)` when the interpreter gives the cell a format, `None` when the file's stands.
    fn formula_format(&mut self, at: Key, formula: &str) -> Option<Code> {
        if !self.busy.insert(at) {
            return None;
        }
        let out = (|| {
            let excel = if self.ods {
                sheetrs_io::reader::ods_parser::ods_formula_to_excel(formula, self.sheet_name(at.0))
            } else {
                formula.to_string()
            };
            let expr = sheetrs_formula::parse(&excel).ok()?;
            if !self.ods && !self.recalculated(at, &expr) {
                return None;
            }
            let val = self.eval(&expr, at);
            let result = self.as_format(&val);
            // An ODS cell without a style takes the default format of the type of its value: a time
            // is `HH:MM:SS AM/PM` and a percentage `0.00%`, whatever the format they come from.
            Some(match (self.ods, class_of(&result)) {
                (true, Class::Time) => code("HH:MM:SS AM/PM"),
                (true, Class::Percent) => code(&Self::percent()),
                _ => result,
            })
        })();
        self.busy.remove(&at);
        out
    }

    /// Whether `LibreOffice` recalculates the cell when it opens an XLSX: it calls a volatile
    /// function, or reads a cell that is recalculated.
    fn recalculated(&mut self, at: Key, expr: &Expr) -> bool {
        if let Some(known) = self.recalculated.get(&at) {
            return *known;
        }
        // A cycle does not make a cell volatile.
        self.recalculated.insert(at, false);
        let mut volatile = false;
        expr.walk(&mut |e| {
            if let Expr::Call(c) = e
                && VOLATILE.contains(&function_name(&c.name).as_str())
            {
                volatile = true;
            }
        });
        if !volatile {
            let refs: Vec<(u16, u32, u32, u32, u32)> = expr
                .references()
                .into_iter()
                .filter_map(|r| self.rect(r, at.0))
                .collect();
            'search: for (sheet, r0, c0, r1, c1) in refs {
                self.ensure(sheet, r0, c0, r1, c1);
                let mut candidates: Vec<(Key, String)> = Vec::new();
                let from = self
                    .cells
                    .partition_point(|(s, r, _, _)| (*s, *r) < (sheet, r0));
                for (s, r, c, cell) in &self.cells[from..] {
                    if *s != sheet || *r > r1 {
                        break;
                    }
                    if (c0..=c1).contains(c)
                        && let Some(f) = &cell.formula
                    {
                        candidates.push(((*s, *r, *c), f.to_string()));
                    }
                }
                for (&(s, r, c), cell) in self.extra.range((sheet, r0, 0)..=(sheet, r1, u32::MAX)) {
                    if s == sheet
                        && (c0..=c1).contains(&c)
                        && let Some(f) = &cell.formula
                    {
                        candidates.push(((s, r, c), f.to_string()));
                    }
                }
                for (key, f) in candidates {
                    if self.depends_on_volatile(key, &f) {
                        volatile = true;
                        break 'search;
                    }
                }
            }
        }
        self.recalculated.insert(at, volatile);
        volatile
    }

    fn depends_on_volatile(&mut self, at: Key, formula: &str) -> bool {
        let excel = if self.ods {
            sheetrs_io::reader::ods_parser::ods_formula_to_excel(formula, self.sheet_name(at.0))
        } else {
            formula.to_string()
        };
        match sheetrs_formula::parse(&excel) {
            Ok(expr) => self.recalculated(at, &expr),
            Err(_) => false,
        }
    }

    /// A reference as a rectangle of the workbook.
    fn rect(&self, r: &Reference, home: u16) -> Option<(u16, u32, u32, u32, u32)> {
        let sheet = match &r.sheet {
            Some(name) => self.sheet_index(name)?,
            None => home,
        };
        Some(match r.area {
            RefArea::Cell(a) => (sheet, a.row, a.col, a.row, a.col),
            RefArea::Range(a, b) => (
                sheet,
                a.row.min(b.row),
                a.col.min(b.col),
                a.row.max(b.row),
                a.col.max(b.col),
            ),
            RefArea::Cols(a, b) => (sheet, 0, a.col.min(b.col), 1_048_575, a.col.max(b.col)),
            RefArea::Rows(a, b) => (sheet, a.row.min(b.row), 0, a.row.max(b.row), 16_383),
        })
    }

    fn as_format(&mut self, v: &Val) -> Code {
        match v {
            Val::Fmt(c) => c.clone(),
            Val::Ref(s, r0, c0, _, _) => self
                .effective((*s, *r0, *c0))
                .filter(|code| !is_text_format(code)),
        }
    }

    /// A constant for an expression that can be worked out without evaluating the workbook.
    fn constant(&mut self, e: &Expr, at: Key) -> Option<Const> {
        match e {
            Expr::Number(n) => Some(Const::Num(n.value)),
            Expr::Text(t) => Some(Const::Text(t.clone())),
            Expr::Bool(b) => Some(Const::Bool(*b)),
            Expr::Unary(UnOp::Neg, x) => Some(Const::Num(-self.constant(x, at)?.number()?)),
            Expr::Paren(x) | Expr::Unary(_, x) => self.constant(x, at),
            Expr::Percent(x) => Some(Const::Num(self.constant(x, at)?.number()? / 100.0)),
            Expr::Ref(r) => {
                let (s, r0, c0, r1, c1) = self.rect(r, at.0)?;
                if (r0, c0) != (r1, c1) {
                    return None;
                }
                self.ensure(s, r0, c0, r0, c0);
                match &self.cell((s, r0, c0))?.value {
                    CellValue::Number(n) => Some(Const::Num(*n)),
                    CellValue::Text(t) => Some(Const::Text(t.to_string())),
                    CellValue::Boolean(b) => Some(Const::Bool(*b)),
                    _ => None,
                }
            }
            Expr::Name(n) => self.through_name(n, |me, expr| me.constant(expr, at)),
            Expr::Binary(op, a, b) => {
                let (x, y) = (self.constant(a, at)?, self.constant(b, at)?);
                Some(match op {
                    BinOp::Add => Const::Num(x.number()? + y.number()?),
                    BinOp::Sub => Const::Num(x.number()? - y.number()?),
                    BinOp::Mul => Const::Num(x.number()? * y.number()?),
                    BinOp::Div => Const::Num(x.number()? / y.number()?),
                    BinOp::Pow => Const::Num(x.number()?.powf(y.number()?)),
                    BinOp::Concat => Const::Text(format!("{}{}", x.text(), y.text())),
                    BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge => {
                        let order = match (&x, &y) {
                            (Const::Text(p), Const::Text(q)) => {
                                p.to_lowercase().cmp(&q.to_lowercase())
                            }
                            _ => x.number()?.partial_cmp(&y.number()?)?,
                        };
                        Const::Bool(match op {
                            BinOp::Eq => order.is_eq(),
                            BinOp::Ne => order.is_ne(),
                            BinOp::Lt => order.is_lt(),
                            BinOp::Le => order.is_le(),
                            BinOp::Gt => order.is_gt(),
                            _ => order.is_ge(),
                        })
                    }
                    BinOp::Range => return None,
                })
            }
            Expr::Call(c) => {
                let name = function_name(&c.name);
                let arg = |i: usize| c.args.get(i).and_then(Option::as_ref);
                match name.as_str() {
                    "TRUE" => Some(Const::Bool(true)),
                    "FALSE" => Some(Const::Bool(false)),
                    "NOT" => Some(Const::Bool(!self.constant(arg(0)?, at)?.truth()?)),
                    "AND" | "OR" => {
                        let mut all = true;
                        let mut any = false;
                        for a in c.args.iter().flatten() {
                            let t = self.constant(a, at)?.truth()?;
                            all &= t;
                            any |= t;
                        }
                        Some(Const::Bool(if name == "AND" { all } else { any }))
                    }
                    _ => self.constant_call(&name, c, at),
                }
            }
            _ => None,
        }
    }

    /// The numbers (a boolean is not one) that the cells of a rectangle hold, as read.
    fn numbers_of(&mut self, rect: (u16, u32, u32, u32, u32)) -> Vec<f64> {
        let (s, r0, c0, r1, c1) = rect;
        self.ensure(s, r0, c0, r1, c1);
        let from = self
            .cells
            .partition_point(|(cs, r, _, _)| (*cs, *r) < (s, r0));
        let mut keys: Vec<Key> = self.cells[from..]
            .iter()
            .take_while(|(cs, r, _, _)| *cs == s && *r <= r1)
            .filter(|(_, _, c, _)| (c0..=c1).contains(c))
            .map(|(cs, r, c, _)| (*cs, *r, *c))
            .collect();
        keys.extend(
            self.extra
                .range((s, r0, 0)..=(s, r1, u32::MAX))
                .filter(|((_, _, c), _)| (c0..=c1).contains(c))
                .map(|(k, _)| *k),
        );
        keys.sort_unstable();
        keys.dedup();
        keys.into_iter()
            .filter_map(|k| match self.cell(k)?.value {
                CellValue::Number(n) => Some(n),
                _ => None,
            })
            .collect()
    }

    /// The numbers of the arguments of `SUM`, `COUNT`, `MAX` and `MIN`: constants and the cells of
    /// ranges.
    fn numbers_of_args(&mut self, c: &Call, at: Key) -> Option<Vec<f64>> {
        let mut out = Vec::new();
        for a in c.args.iter().flatten() {
            if let Expr::Ref(r) = a {
                let rect = self.rect(r, at.0)?;
                out.extend(self.numbers_of(rect));
            } else if let Some(Const::Num(n)) = self.constant(a, at) {
                out.push(n);
            } else {
                return None;
            }
        }
        Some(out)
    }

    /// The functions that give a constant when their arguments do (the ones the formulas that choose
    /// a branch or a cell use to say which).
    fn constant_call(&mut self, name: &str, c: &Call, at: Key) -> Option<Const> {
        let arg = |i: usize| c.args.get(i).and_then(Option::as_ref);
        let number = |me: &mut Self, i: usize| me.constant(arg(i)?, at)?.number();
        match name {
            "IF" => {
                let branch = if self.constant(arg(0)?, at)?.truth()? {
                    arg(1)?
                } else {
                    arg(2)?
                };
                self.constant(branch, at)
            }
            "ISNUMBER" | "ISTEXT" => match self.constant(arg(0)?, at)? {
                Const::Num(_) => Some(Const::Bool(name == "ISNUMBER")),
                Const::Text(_) => Some(Const::Bool(name == "ISTEXT")),
                // `LibreOffice` and `Excel` do not agree on a boolean.
                Const::Bool(_) => None,
            },
            "ISBLANK" => {
                let Expr::Ref(r) = arg(0)? else { return None };
                let (s, r0, c0, r1, c1) = self.rect(r, at.0)?;
                if (r0, c0) != (r1, c1) {
                    return None;
                }
                self.ensure(s, r0, c0, r0, c0);
                Some(Const::Bool(self.cell((s, r0, c0)).is_none_or(|cell| {
                    matches!(cell.value, CellValue::Empty) && cell.formula.is_none()
                })))
            }
            "LEN" => Some(Const::Num(
                self.constant(arg(0)?, at)?.text().chars().count() as f64,
            )),
            "ABS" => Some(Const::Num(number(self, 0)?.abs())),
            "INT" => Some(Const::Num(number(self, 0)?.floor())),
            "ROUND" => {
                let (x, digits) = (number(self, 0)?, number(self, 1).unwrap_or(0.0));
                let scale = 10f64.powi(digits as i32);
                Some(Const::Num((x * scale).round() / scale))
            }
            "MOD" => {
                let (x, y) = (number(self, 0)?, number(self, 1)?);
                (y != 0.0).then(|| Const::Num(x - y * (x / y).floor()))
            }
            "ROW" | "COLUMN" => {
                let (row, col) = match arg(0) {
                    None => (at.1, at.2),
                    Some(Expr::Ref(r)) => {
                        let (_, r0, c0, _, _) = self.rect(r, at.0)?;
                        (r0, c0)
                    }
                    Some(_) => return None,
                };
                Some(Const::Num(
                    f64::from(if name == "ROW" { row } else { col }) + 1.0,
                ))
            }
            "CONCATENATE" | "CONCAT" => {
                let mut text = String::new();
                for a in c.args.iter().flatten() {
                    text.push_str(&self.constant(a, at)?.text());
                }
                Some(Const::Text(text))
            }
            "UPPER" => Some(Const::Text(
                self.constant(arg(0)?, at)?.text().to_uppercase(),
            )),
            "LOWER" => Some(Const::Text(
                self.constant(arg(0)?, at)?.text().to_lowercase(),
            )),
            "LEFT" | "RIGHT" => {
                let text: Vec<char> = self.constant(arg(0)?, at)?.text().chars().collect();
                let n = if arg(1).is_some() {
                    number(self, 1)?
                } else {
                    1.0
                };
                if n < 0.0 {
                    return None;
                }
                let n = (n as usize).min(text.len());
                Some(Const::Text(if name == "LEFT" {
                    text[..n].iter().collect()
                } else {
                    text[text.len() - n..].iter().collect()
                }))
            }
            "ADDRESS" if c.args.len() == 2 => {
                let (row, col) = (number(self, 0)?, number(self, 1)?);
                if row < 1.0 || col < 1.0 {
                    return None;
                }
                Some(Const::Text(format!(
                    "${}${}",
                    sheetrs_formula::addr::col_name(col as u32 - 1),
                    row as u32
                )))
            }
            "SUM" | "COUNT" | "MAX" | "MIN" => {
                let numbers = self.numbers_of_args(c, at)?;
                Some(Const::Num(match name {
                    "SUM" => numbers.iter().sum(),
                    "COUNT" => numbers.len() as f64,
                    "MAX" => numbers.iter().copied().reduce(f64::max).unwrap_or(0.0),
                    _ => numbers.iter().copied().reduce(f64::min).unwrap_or(0.0),
                }))
            }
            _ => None,
        }
    }

    fn eval(&mut self, e: &Expr, at: Key) -> Val {
        match e {
            Expr::Number(_)
            | Expr::Text(_)
            | Expr::Error(_)
            | Expr::Array(_)
            | Expr::Unary(UnOp::Neg, _) => Val::Fmt(None),
            Expr::Name(n) => self
                .through_name(n, |me, expr| Some(me.eval(expr, at)))
                .unwrap_or(Val::Fmt(None)),
            Expr::Bool(_) => Val::Fmt(code("BOOLEAN")),
            Expr::Paren(x) | Expr::Unary(UnOp::Plus | UnOp::Implicit, x) => self.eval(x, at),
            Expr::Percent(x) => {
                let inner = self.eval(x, at);
                let c = self.as_format(&inner);
                Val::Fmt(if class_of(&c) == Class::Currency {
                    c
                } else {
                    code(&Self::percent())
                })
            }
            Expr::Ref(r) => match self.rect(r, at.0) {
                Some((s, r0, c0, r1, c1)) => {
                    self.ensure(s, r0, c0, r1, c1);
                    Val::Ref(s, r0, c0, r1, c1)
                }
                None => Val::Fmt(None),
            },
            Expr::Binary(op, a, b) => {
                let (x, y) = (self.eval(a, at), self.eval(b, at));
                let (x, y) = (self.as_format(&x), self.as_format(&y));
                Val::Fmt(self.binary(*op, &x, &y))
            }
            Expr::Call(c) => self.call(c, at),
        }
    }

    fn percent() -> String {
        "0.00%".to_string()
    }

    /// The default format of a date, a date and a time, a time and a duration, as the locale of the
    /// text context writes them.
    fn default_date(&self) -> Code {
        code(&self.text.default_date_format())
    }

    fn default_date_time(&self) -> Code {
        code(&format!("{} HH:MM AM/PM", self.text.default_date_format()))
    }

    fn binary(&self, op: BinOp, a: &Code, b: &Code) -> Code {
        use BinOp::{Add, Div, Mul, Sub};
        let (ca, cb) = (class_of(a), class_of(b));
        match op {
            Add | Sub => self.add(a, b, ca, cb),
            Mul => {
                if ca == Class::Currency {
                    a.clone()
                } else if cb == Class::Currency {
                    b.clone()
                } else {
                    None
                }
            }
            Div => {
                if ca == Class::Currency && cb != Class::Currency {
                    a.clone()
                } else {
                    None
                }
            }
            BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge => {
                code("BOOLEAN")
            }
            _ => None,
        }
    }

    /// `+` and `-` by the class of the operands (rows are the first operand).
    ///
    /// | | date | date+time | time | percent | currency | number | plain |
    /// |-|------|-----------|------|---------|----------|--------|-------|
    /// | **date** | General | General | date+time | date | currency | number | date |
    /// | **date+time** | duration | duration | date+time | date+time | currency | number | date+time |
    /// | **time** | date+time | date+time | duration | time | currency | number | time |
    /// | **percent** | date | date+time | time | percent | currency | percent | percent |
    /// | **currency** | currency | currency | currency | currency | currency | currency | currency |
    /// | **number** | number | number | number | percent | currency | General | General |
    /// | **plain** | date | date+time | time | percent | currency | General | General |
    ///
    /// (`date+time` with `date+time` is the duration; "currency", "number" copy the operand's format.)
    fn add(&self, a: &Code, b: &Code, ca: Class, cb: Class) -> Code {
        use Class::{Boolean, Currency, Date, DateTime, Number, Percent, Plain, Time};
        let plain = |c: Class| {
            if matches!(c, Boolean | Class::Scientific) {
                Plain
            } else {
                c
            }
        };
        let (ca, cb) = (plain(ca), plain(cb));
        let duration = || code("[HH]:MM:SS");
        let time = || code("HH:MM:SS AM/PM");
        let pct = || code(&Self::percent());
        if ca == Currency && cb == Currency {
            return a.clone();
        }
        if ca == Currency {
            return a.clone();
        }
        if cb == Currency {
            return b.clone();
        }
        match (ca, cb) {
            (Date | DateTime | Time, Number) => b.clone(),
            (Number, Date | DateTime | Time) => a.clone(),
            (Date, Percent | Plain) | (Percent | Plain, Date) => self.default_date(),
            (Date, Time)
            | (DateTime, Time | Percent | Plain)
            | (Time, Date | DateTime)
            | (Percent | Plain, DateTime) => self.default_date_time(),
            (DateTime, DateTime) | (Time, Time) => duration(),
            (Time, Percent | Plain) | (Percent | Plain, Time) => time(),
            (Percent, Percent | Number | Plain) | (Number | Plain, Percent) => pct(),
            // Date with date, number or plain with each other, and the rest: `General`.
            _ => None,
        }
    }

    /// The format an aggregate takes from the cell `cell`: its own for most, the default percentage
    /// for a percentage, `General` for a number format.
    fn aggregate(&mut self, cell: Option<Key>, keeps_boolean: bool) -> Code {
        let key = cell?;
        let c = self.effective(key).filter(|code| !is_text_format(code));
        match class_of(&c) {
            Class::Percent => code(&Self::percent()),
            Class::Scientific => code("0.00E+00"),
            Class::Number => None,
            Class::Boolean => {
                if keeps_boolean {
                    c
                } else {
                    None
                }
            }
            _ => c,
        }
    }

    fn call(&mut self, c: &Call, at: Key) -> Val {
        let name = function_name(&c.name);
        let arg = |i: usize| c.args.get(i).and_then(Option::as_ref);
        let general = Val::Fmt(None);
        match name.as_str() {
            "INDIRECT" => {
                let Some(text) = arg(0).and_then(|a| self.constant(a, at)) else {
                    return general;
                };
                let text = text.text();
                let (sheet, tail) = split_sheet(text.trim());
                let Some(area) = parse_area(tail) else {
                    // The text can be the name of a range.
                    return match sheet {
                        None => self
                            .through_name(text.trim(), |me, expr| Some(me.eval(expr, at)))
                            .unwrap_or(general),
                        Some(_) => general,
                    };
                };
                let sheet = match sheet {
                    Some(n) => match self.sheet_index(&n) {
                        Some(i) => i,
                        None => return general,
                    },
                    None => at.0,
                };
                self.ensure(sheet, area.r0, area.c0, area.r1, area.c1);
                Val::Ref(sheet, area.r0, area.c0, area.r1, area.c1)
            }
            "OFFSET" => {
                let Some(Val::Ref(s, r0, c0, r1, c1)) = arg(0).map(|a| self.eval(a, at)) else {
                    return general;
                };
                let n = |me: &mut Self, i: usize| -> Option<i64> {
                    me.constant(arg(i)?, at)?.number().map(|x| x as i64)
                };
                let (Some(dr), Some(dc)) = (n(self, 1), n(self, 2)) else {
                    return general;
                };
                let height = n(self, 3).unwrap_or(i64::from(r1 - r0) + 1);
                let width = n(self, 4).unwrap_or(i64::from(c1 - c0) + 1);
                let (nr, nc) = (i64::from(r0) + dr, i64::from(c0) + dc);
                if nr < 0 || nc < 0 || height < 1 || width < 1 {
                    return general;
                }
                let (nr, nc) = (nr as u32, nc as u32);
                let (er, ec) = (nr + height as u32 - 1, nc + width as u32 - 1);
                self.ensure(s, nr, nc, er, ec);
                Val::Ref(s, nr, nc, er, ec)
            }
            "INDEX" => {
                let Some(Val::Ref(s, r0, c0, r1, c1)) = arg(0).map(|a| self.eval(a, at)) else {
                    return general;
                };
                let n = |me: &mut Self, i: usize| -> Option<i64> {
                    me.constant(arg(i)?, at)?.number().map(|x| x as i64)
                };
                let rows = i64::from(r1 - r0) + 1;
                let cols = i64::from(c1 - c0) + 1;
                // One column or one row: the single index runs along it.
                let (row, col) = match (n(self, 1), n(self, 2)) {
                    (Some(r), Some(c)) => (r, c),
                    (Some(i), None) if cols == 1 => (i, 1),
                    (Some(i), None) if rows == 1 => (1, i),
                    (Some(i), None) => (i, 1),
                    _ => return general,
                };
                if row < 1 || col < 1 || row > rows || col > cols {
                    return general;
                }
                let (rr, cc) = (r0 + row as u32 - 1, c0 + col as u32 - 1);
                Val::Ref(s, rr, cc, rr, cc)
            }
            "CHOOSE" => {
                let Some(i) = arg(0)
                    .and_then(|a| self.constant(a, at))
                    .and_then(|k| k.number())
                else {
                    return general;
                };
                match arg(i as usize) {
                    Some(branch) if i >= 1.0 => self.eval(branch, at),
                    _ => general,
                }
            }
            "IF" => {
                let cond = arg(0)
                    .and_then(|a| self.constant(a, at))
                    .and_then(|k| k.truth());
                let branch = |me: &mut Self, i: usize, missing_is_false: bool| -> Val {
                    match arg(i) {
                        Some(e) => me.eval(e, at),
                        None if missing_is_false => Val::Fmt(code("BOOLEAN")),
                        None => Val::Fmt(None),
                    }
                };
                match cond {
                    Some(true) => branch(self, 1, false),
                    Some(false) => branch(self, 2, true),
                    None => {
                        // The two branches give the same format: nothing to choose.
                        let (a, b) = (branch(self, 1, false), branch(self, 2, true));
                        let (fa, fb) = (self.as_format(&a), self.as_format(&b));
                        if fa == fb { Val::Fmt(fa) } else { general }
                    }
                }
            }
            "IFERROR" | "IFNA" => arg(0).map_or(general, |a| self.eval(a, at)),
            "SUM" | "MIN" | "MAX" | "AVERAGE" | "PRODUCT" | "SUMSQ" | "AVERAGEA" | "MAXA"
            | "MINA" | "SUBTOTAL" | "AGGREGATE" => {
                let first = usize::from(matches!(name.as_str(), "SUBTOTAL" | "AGGREGATE"));
                let first = if name == "AGGREGATE" { 2 } else { first };
                let last = name == "SUM";
                let keeps_boolean = matches!(name.as_str(), "MIN" | "MAX" | "MINA" | "MAXA");
                // The arguments are looked at in order and the first one that holds a number
                // gives the format: the cells that hold none (text, empty) and the arguments after
                // it do not count.
                for a in c.args.iter().skip(first).flatten() {
                    // `SUM(C:C)` over a whole column or row is `General` whichever cells it holds.
                    let whole = matches!(
                        a,
                        Expr::Ref(Reference {
                            area: RefArea::Cols(..) | RefArea::Rows(..),
                            ..
                        })
                    );
                    if last && whole {
                        return general;
                    }
                    match self.eval(a, at) {
                        Val::Ref(s, r0, c0, r1, c1) => {
                            let cells = self.numbers_in(s, r0, c0, r1, c1);
                            let cell = if last { cells.last() } else { cells.first() };
                            if let Some(&key) = cell {
                                return Val::Fmt(self.aggregate(Some(key), keeps_boolean));
                            }
                        }
                        // A value that is not a reference: its format as the rules of operators say.
                        Val::Fmt(f) => {
                            return Val::Fmt(match class_of(&f) {
                                Class::Percent => code(&Self::percent()),
                                Class::Scientific => code("0.00E+00"),
                                Class::Number | Class::Boolean => None,
                                _ => f,
                            });
                        }
                    }
                }
                general
            }
            "TRUE" | "FALSE" | "NOT" | "AND" | "OR" | "XOR" | "EXACT" => Val::Fmt(code("BOOLEAN")),
            n if n.starts_with("IS") && n != "ISO" => Val::Fmt(code("BOOLEAN")),
            "DATE" | "TODAY" => Val::Fmt(self.default_date()),
            "NOW" => Val::Fmt(self.default_date_time()),
            "TIME" => Val::Fmt(code("HH:MM:SS AM/PM")),
            _ => general,
        }
    }
}

fn is_number(cell: &Cell) -> bool {
    matches!(cell.value, CellValue::Number(_) | CellValue::Boolean(_))
}

fn own_is_general(cell: &Cell) -> bool {
    own_code(cell).is_none()
}

/// The format a cell carries in the file, `None` for `General`.
fn own_code(cell: &Cell) -> Code {
    let f = cell.num_fmt.as_deref()?;
    let mut rest = f;
    while let Some(after) = rest.strip_prefix("[$-")
        && let Some((_, tail)) = after.split_once(']')
    {
        rest = tail;
    }
    if rest.eq_ignore_ascii_case("general") {
        None
    } else {
        Some(Rc::from(f))
    }
}
