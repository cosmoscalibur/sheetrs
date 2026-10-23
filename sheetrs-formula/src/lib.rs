//! Lossless spreadsheet-formula AST: a parser, a canonical printer and
//! reference rewriting for structural edits.
//!
//! The AST keeps absolute markers, function names as written, number text,
//! parentheses and omitted arguments, so a formula can be analysed, rewritten
//! (insert/delete rows and columns, rename or delete a sheet) and printed back
//! without losing meaning. Whitespace is not preserved: printing gives the
//! canonical spelling and `parse(print(parse(s))) == parse(s)`.

pub mod addr;
pub mod ast;
pub mod edit;
pub mod parse;
pub mod print;
pub mod r1c1;

pub use ast::Expr;
pub use parse::{ParseError, parse};
