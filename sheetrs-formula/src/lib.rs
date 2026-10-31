//! Lossless spreadsheet-formula AST: a parser, a canonical printer and
//! reference rewriting for structural edits.
//!
//! The AST keeps absolute markers, function names as written, number text,
//! parentheses and omitted arguments, so a formula can be analysed, rewritten
//! (insert/delete rows and columns, rename or delete a sheet) and printed back
//! without losing meaning. Printing gives the canonical spelling and
//! `parse(print(parse(s))) == parse(s)`. Whitespace and spelling are not in the
//! AST but are not lost: [`parse_spanned`] keeps the source and the span of each
//! reference, so rules can inspect the text between tokens and
//! [`Formula::rewrite`] changes only the references an edit touches, leaving the
//! rest of the text byte for byte as written.

pub mod addr;
pub mod ast;
pub mod calc_print;
pub mod edit;
pub mod grid;
pub mod parse;
pub mod print;
pub mod r1c1;

pub use ast::Expr;
pub use parse::{Formula, ParseError, parse, parse_spanned};
