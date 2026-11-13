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
//!
//! ```
//! use sheetrs_formula::parse;
//!
//! let formula = parse("SUM(A1:B2)+1").expect("a valid formula");
//! // Printing gives the canonical spelling, and parsing it again gives the same formula.
//! assert_eq!(parse(&formula.to_string()).expect("a valid formula"), formula);
//! ```

pub mod addr;
pub mod ast;
pub mod calc_print;
pub mod calc_read;
pub mod calendar;
mod date_system;
pub mod edit;
pub mod grid;
pub mod lax;
pub mod parse;
pub mod print;

pub use ast::Expr;
pub use date_system::{DateSystem, ParseDateSystemError, UnsupportedDateSystem};
pub use parse::{Formula, ParseError, parse, parse_spanned};
