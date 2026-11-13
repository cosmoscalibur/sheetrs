//! Format-independent intermediate representation.
//!
//! Neither XLSX nor ODS survive past the frontend: formulas are lowered to
//! [`Expr`], cells to [`Program`] slots, and a single engine executes them.
//! Functions are identified by a namespaced [`FuncId`] (`core::SUM`,
//! `ods::EASTERSUNDAY`) so that behavioral differences between the formats can
//! coexist in one registry.

use crate::addr::{CellRef, Rect};
use crate::functions;
use crate::options::EngineOptions;
use crate::value::Value;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

/// Namespace of a function, as `OpenDocument` names them. `Core` is the standard: its names carry
/// no prefix (`SUM`). The others are the prefixes a name carries in an ODS and, with `_xlfn.`,
/// in an XLSX: `COM.MICROSOFT.` (functions Microsoft defined after ODF 1.2, and Excel's own
/// version of a function whose bare name is something else), `ORG.OPENOFFICE.` and
/// `ORG.LIBREOFFICE.` (extensions of the two projects) and `LEGACY.` (the pre-ODF-1.2 version of
/// a function whose bare name now means another).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Ns {
    Core,
    Microsoft,
    OpenOffice,
    LibreOffice,
    Legacy,
}

impl Ns {
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Core => "core",
            Self::Microsoft => "microsoft",
            Self::OpenOffice => "openoffice",
            Self::LibreOffice => "libreoffice",
            Self::Legacy => "legacy",
        }
    }

    #[must_use]
    pub fn parse(s: &str) -> Option<Self> {
        Some(match s {
            "core" => Self::Core,
            "microsoft" => Self::Microsoft,
            "openoffice" => Self::OpenOffice,
            "libreoffice" => Self::LibreOffice,
            "legacy" => Self::Legacy,
            _ => return None,
        })
    }

    /// The namespace an ODS name's prefix names (`COM.MICROSOFT.` …), with the name that is left.
    #[must_use]
    pub fn of_prefixed(name: &str) -> Option<(Self, &str)> {
        [
            ("COM.MICROSOFT.", Self::Microsoft),
            ("ORG.OPENOFFICE.", Self::OpenOffice),
            ("ORG.LIBREOFFICE.", Self::LibreOffice),
            ("LEGACY.", Self::Legacy),
        ]
        .into_iter()
        .find_map(|(p, ns)| name.strip_prefix(p).map(|rest| (ns, rest)))
    }
}

/// Index into the function registry. Serialized as `"ns::NAME"` so programs stay valid when
/// the registry grows.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FuncId(pub u16);

impl FuncId {
    #[must_use]
    pub fn def(self) -> &'static functions::FuncDef {
        &functions::table()[self.0 as usize]
    }

    #[must_use]
    pub fn qualified(self) -> String {
        let d = self.def();
        format!("{}::{}", d.ns.as_str(), d.name)
    }
}

impl Serialize for FuncId {
    fn serialize<S: Serializer>(&self, s: S) -> Result<S::Ok, S::Error> {
        s.serialize_str(&self.qualified())
    }
}

impl<'de> Deserialize<'de> for FuncId {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        let (ns, name) = s
            .split_once("::")
            .and_then(|(ns, n)| Ns::parse(ns).map(|ns| (ns, n)))
            .ok_or_else(|| serde::de::Error::custom(format!("bad function id `{s}`")))?;
        functions::lookup_in(ns, name)
            .ok_or_else(|| serde::de::Error::custom(format!("unknown function `{s}`")))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum UnOp {
    Neg,
    Plus,
    Percent,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Expr {
    Const(Value),
    Cell(CellRef),
    Range(Rect),
    Unary(UnOp, Box<Expr>),
    Binary(BinOp, Box<Expr>, Box<Expr>),
    Call(FuncId, Vec<Expr>),
    /// Call to a function missing from the registry, by name and arguments: evaluates to
    /// `#NAME?`. Boxed because it is rare and would otherwise set the size of every node.
    Unknown(Box<(Box<str>, Vec<Expr>)>),
    /// Omitted argument (`IF(A1,,2)`).
    Missing,
}

impl Expr {
    /// Visit this node and all descendants.
    pub fn walk(&self, f: &mut impl FnMut(&Expr)) {
        f(self);
        match self {
            Self::Unary(_, a) => a.walk(f),
            Self::Binary(_, a, b) => {
                a.walk(f);
                b.walk(f);
            }
            Self::Call(_, args) => args.iter().for_each(|a| a.walk(f)),
            Self::Unknown(call) => call.1.iter().for_each(|a| a.walk(f)),
            _ => {}
        }
    }

    pub fn walk_mut(&mut self, f: &mut impl FnMut(&mut Expr)) {
        match self {
            Self::Unary(_, a) => a.walk_mut(f),
            Self::Binary(_, a, b) => {
                a.walk_mut(f);
                b.walk_mut(f);
            }
            Self::Call(_, args) => args.iter_mut().for_each(|a| a.walk_mut(f)),
            Self::Unknown(call) => call.1.iter_mut().for_each(|a| a.walk_mut(f)),
            _ => {}
        }
        f(self);
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum CellKind {
    Const(Value),
    Formula(Expr),
    /// Run-time parameter, with the value found in the source as default.
    Input(Value),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellDef {
    pub at: CellRef,
    pub kind: CellKind,
}

/// A named group of cells exchanged with the outside world as JSON.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Binding {
    pub name: String,
    pub rect: Rect,
}

/// The compiled artifact. `cells` is in evaluation order: dependencies first.
/// Slots listed in `cycles` (half-open ranges) form strongly connected
/// components that need iterative calculation.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Program {
    pub sheets: Vec<String>,
    pub options: EngineOptions,
    pub cells: Vec<CellDef>,
    /// Hidden rows `(sheet, first, last)`, 0-based inclusive: `SUBTOTAL(109…)` and
    /// `AGGREGATE` skip them.
    #[serde(default)]
    pub hidden_rows: Vec<(u16, u32, u32)>,
    /// Cells whose formula holds `SUBTOTAL` or `AGGREGATE` (sorted): a subtotal
    /// does not count them again.
    #[serde(default)]
    pub subtotal_cells: Vec<CellRef>,
    pub cycles: Vec<(u32, u32)>,
    pub inputs: Vec<Binding>,
    pub outputs: Vec<Binding>,
    /// The names, besides the bindings, a caller may use when running.
    pub queryable: Queryable,
    /// What the program is for: values in and out, or the workbook itself as the answer.
    #[serde(default)]
    pub mode: Mode,
    /// The formula cells the workbook answer writes, as runs in the order of their positions: in a
    /// workbook mode (`Mode::is_workbook`) the formulas inside the outputs (all of them when no
    /// output is declared), and none otherwise. The other formulas of the workbook keep the result
    /// the template stored.
    ///
    /// The workbook is its own template, so the program is what knows where its formulas are: the
    /// writer of the answer takes the positions to write from here and does not discover them in the
    /// file. A formula the compiler folded, kept as the value the file cached (the parser did not
    /// take it), or that lies inside an array formula is in the list all the same.
    #[serde(default)]
    pub formula_cells: Vec<CellRun>,
    /// Entries of the workbook's settings sheet, inputs and outputs included.
    pub meta: Vec<Meta>,
    pub diagnostics: Vec<String>,
    /// How and from what the program was built.
    #[serde(default)]
    pub build: BuildInfo,
}

/// A run of consecutive cells of one row: `first..=last` columns of `row` in `sheet` (0-based).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct CellRun {
    pub sheet: u16,
    pub row: u32,
    pub first: u32,
    pub last: u32,
}

impl CellRun {
    /// The runs of the positions `cells`, which must come in the order of their positions.
    pub fn runs(cells: impl IntoIterator<Item = CellRef>) -> Vec<Self> {
        let mut runs: Vec<Self> = Vec::new();
        for at in cells {
            match runs.last_mut() {
                Some(r) if r.sheet == at.sheet && r.row == at.row && r.last + 1 == at.col => {
                    r.last = at.col;
                }
                _ => runs.push(Self {
                    sheet: at.sheet,
                    row: at.row,
                    first: at.col,
                    last: at.col,
                }),
            }
        }
        runs
    }

    /// The positions of the run.
    pub fn cells(&self) -> impl Iterator<Item = CellRef> + '_ {
        (self.first..=self.last).map(|col| CellRef::new(self.sheet, self.row, col))
    }
}

/// How and from what a program was built, listed by `--info`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct BuildInfo {
    /// The compiler, with its version (`sheetrs-compiler 0.3.2`).
    pub compiler: String,
    /// The platform of the compiler (`x86_64-linux`).
    pub target: String,
    /// When the program was built, in UTC (`2026-10-30T12:00:00Z`); `SOURCE_DATE_EPOCH` fixes it.
    pub built_at: String,
    pub options: BuildOptions,
    pub source: SourceInfo,
}

/// The build options that shape the program.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct BuildOptions {
    /// `declared` (the outputs were declared) or `deduced` (the formula cells, as the rectangles that hold them).
    pub outputs: String,
    /// `declared` (the inputs were declared) or `deduced` (the constants the outputs start from).
    #[serde(default)]
    pub inputs: String,
    /// The locale tag the program is fixed to.
    pub locale: String,
    /// Whether input-independent cells were evaluated at build time.
    pub fold: bool,
    /// The compatibility features that were on.
    pub features: Vec<String>,
}

/// The workbook a program was built from.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct SourceInfo {
    /// The file name, with its extension.
    pub file: String,
    /// `xlsx` or `ods`.
    pub format: String,
    /// SHA-256 of the file, in hexadecimal.
    pub sha256: String,
}

/// The names a request may use besides the bindings: the defined names and tables of the
/// workbook that point inside what the program exposes. A name inside the outputs is an alias of
/// that output; a name inside the inputs is an alias of those inputs.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct Queryable {
    pub names: Vec<Binding>,
}

/// What the program is for.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum Mode {
    /// Values in, values out: the inputs and outputs of `Program::inputs` and `Program::outputs`.
    #[default]
    Values,
    /// Values in, values out as in any program, and the workbook embedded as a template: the answer
    /// can also be the workbook with the outputs that are formulas, and the inputs given, written into
    /// it (`--output`), the JSON, or both. Inputs and outputs are declared or deduced as in any
    /// program; the formulas outside the program keep the result the template stored.
    Embedded,
    /// The editor's recalculation: as `Embedded`, with every formula an output and the declarations of the
    /// workbook ignored, to bring the stale results of a workbook up to date before an edit reads them.
    /// No command and no option builds it: what a user asks for is `calculate`, with inputs and
    /// outputs declared or deduced.
    Recalculate,
}

impl Mode {
    /// Whether the answer is the workbook.
    #[must_use]
    pub fn is_workbook(self) -> bool {
        self != Self::Values
    }
}

/// A JSON value in a binary program: as its text (a binary format cannot hold a
/// self-describing value); JSON files keep it as a value.
mod json_text {
    use serde::{Deserialize, Deserializer, Serialize, Serializer};

    pub(super) fn serialize<S: Serializer>(v: &serde_json::Value, s: S) -> Result<S::Ok, S::Error> {
        if s.is_human_readable() {
            v.serialize(s)
        } else {
            v.to_string().serialize(s)
        }
    }

    pub(super) fn deserialize<'de, D: Deserializer<'de>>(
        d: D,
    ) -> Result<serde_json::Value, D::Error> {
        if d.is_human_readable() {
            serde_json::Value::deserialize(d)
        } else {
            let text = String::deserialize(d)?;
            serde_json::from_str(&text).map_err(serde::de::Error::custom)
        }
    }
}

/// One row of the settings sheet.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Meta {
    pub key: String,
    #[serde(with = "json_text")]
    pub value: serde_json::Value,
    pub description: String,
    /// `input`, `output` or a custom type.
    pub meta_type: String,
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A program holds one `Expr` per node of every formula: its size is the memory of the build.
    #[cfg(target_pointer_width = "64")]
    #[test]
    fn expr_node_is_32_bytes() {
        assert_eq!(std::mem::size_of::<Expr>(), 32);
    }

    /// A call to a function the registry lacks keeps the form it was written in, in JSON and in
    /// the binary program.
    #[test]
    fn unknown_call_round_trips_in_both_encodings() {
        let call = Expr::Unknown(Box::new((
            "NOPE".into(),
            vec![Expr::Const(Value::Number(1.0)), Expr::Missing],
        )));
        let json = serde_json::to_string(&call).unwrap();
        assert_eq!(
            json,
            r#"{"Unknown":["NOPE",[{"Const":{"Number":1.0}},"Missing"]]}"#
        );
        assert_eq!(serde_json::from_str::<Expr>(&json).unwrap(), call);
        let bytes = postcard::to_allocvec(&call).unwrap();
        assert_eq!(postcard::from_bytes::<Expr>(&bytes).unwrap(), call);
    }
}
