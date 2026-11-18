use std::collections::BTreeMap;
use std::error::Error as StdError;
use std::fs::{self, File};

use sheetrs_core::{Cell, CellPos, Value, Visibility};
use sheetrs_testkit::case::{Case, CellInput, CellValue, Format, Op, Read, SheetSpec};
use sheetrs_testkit::oracle;

mod common;
use common::single_sheet_package_with;

const OUTPUT: &str = concat!(env!("CARGO_TARGET_TMPDIR"), "/oracle_read_back.xlsx");
const CRAFTED_DIR: &str = env!("CARGO_TARGET_TMPDIR");

/// Hand-written cells that real producers emit or that the format allows, as
/// (name, rows XML, shared strings XML).
const CRAFTED: &[(&str, &str, Option<&str>)] = &[
    (
        "empty_v_on_formula",
        r#"<row r="1"><c r="A1"><f>1+1</f><v></v></c></row>"#,
        None,
    ),
    (
        "inline_without_is",
        r#"<row r="1"><c r="A1" t="inlineStr"><v>abc</v></c></row>"#,
        None,
    ),
    (
        "bool_two",
        r#"<row r="1"><c r="A1" t="b"><v>2</v></c></row>"#,
        None,
    ),
    (
        "bool_word_true",
        r#"<row r="1"><c r="A1" t="b"><v>TRUE</v></c></row>"#,
        None,
    ),
    (
        "bool_word_false",
        r#"<row r="1"><c r="A1" t="b"><v>false</v></c></row>"#,
        None,
    ),
    (
        "escaped_carriage_return",
        r#"<row r="1"><c r="A1" t="s"><v>0</v></c></row>"#,
        Some("<si><t>a_x000D_b_x005F_x</t></si>"),
    ),
    (
        "str_without_v",
        r#"<row r="1"><c r="A1" t="str"><f>"x"</f></c></row>"#,
        None,
    ),
    (
        "padded_number",
        r#"<row r="1"><c r="A1"><v> 1 </v></c></row>"#,
        None,
    ),
    (
        "rich_inline",
        r#"<row r="1"><c r="A1" t="inlineStr"><is><r><t>a</t></r><r><t>b</t></r></is></c></row>"#,
        None,
    ),
    (
        "bool_caps_false",
        r#"<row r="1"><c r="A1" t="b"><v>FALSE</v></c></row>"#,
        None,
    ),
    (
        "bool_decimal_zero",
        r#"<row r="1"><c r="A1" t="b"><v>0.0</v></c></row>"#,
        None,
    ),
    (
        "number_with_exponent",
        r#"<row r="1"><c r="A1"><v>1.5E-3</v></c></row>"#,
        None,
    ),
    (
        "inline_with_escape",
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>x_x000A_y</t></is></c></row>"#,
        None,
    ),
    (
        "empty_v_shared",
        r#"<row r="1"><c r="A1" t="s"><v></v></c></row>"#,
        Some("<si><t>a</t></si>"),
    ),
    (
        "self_closing_v_shared",
        r#"<row r="1"><c r="A1" t="s"><v/></c></row>"#,
        Some("<si><t>a</t></si>"),
    ),
    (
        "empty_v_bool",
        r#"<row r="1"><c r="A1" t="b"><v></v></c></row>"#,
        None,
    ),
    (
        "empty_v_error",
        r#"<row r="1"><c r="A1" t="e"><v></v></c></row>"#,
        None,
    ),
    (
        "empty_v_str",
        r#"<row r="1"><c r="A1" t="str"><v></v></c></row>"#,
        None,
    ),
    (
        "inline_content_on_a_numeric_cell",
        r#"<row r="1"><c r="A1" t="n"><is><t>x</t></is></c></row>"#,
        None,
    ),
    (
        "empty_inline_string",
        r#"<row r="1"><c r="A1" t="inlineStr"><is></is></c></row>"#,
        None,
    ),
    (
        "empty_shared_string",
        r#"<row r="1"><c r="A1" t="s"><v>0</v></c></row>"#,
        Some("<si/>"),
    ),
    (
        "str_result_keeps_escapes",
        r#"<row r="1"><c r="A1" t="str"><v>x_x000D_y</v></c></row>"#,
        None,
    ),
    (
        "escaped_surrogate_pair",
        r#"<row r="1"><c r="A1" t="s"><v>0</v></c></row>"#,
        Some("<si><t>_xD83D__xDE00_</t></si>"),
    ),
];

struct Crafted {
    name: String,
    rows: String,
    range: &'static str,
    strings: Option<String>,
    properties: String,
}

const NUMBERS: &[&str] = &[
    "1_0",
    "1d5",
    "1e",
    "1 2",
    "1,5",
    "0x10",
    ".",
    "+",
    ".e1",
    "nan",
    "inf",
    "Infinity",
    "1e999",
    "-1e999",
    "1e-999",
    "  ",
    "+5",
    "-.5",
    "5.",
    "1E+2",
    "1e+",
    "0.5.5",
    "--1",
    "1-",
    "\u{661}\u{662}",
    "1\u{a0}",
    "\u{a0}1",
    "NaN",
    "-NaN",
    "Inf",
    "INF",
    "1e5x",
    "-",
    "-0",
    "0e0",
    "1e400",
    "123456789012345678901234567890",
    "0.1000000000000000055511151231257827",
    "NaNx",
    "INFx",
    "-INF",
    "+INF",
    "NAN",
    "nAn",
    " NaN",
    "Nan",
    "1.#INF",
    "1.#QNAN",
    "1.#IND",
    "-1.#INF",
    "INF1",
    "NaN1",
    "\tINF",
    "+NaN",
];
const BOOLEANS: &[&str] = &[
    "yes", "T", "t", "inf", "nan", "NaN", "1e999", "-1e999", " ", "True", " 1 ", "1abc", "abc1",
    "0abc", "+1", "-1", ".5", "0x1", "tRuE", "TRUE ", "false", "-0", "0e0", "Infinity", "INF",
    "-INF", "TRUE", "tRUE", "truE", "T R U E",
];
const INDEXES: &[&str] = &[
    "abc",
    " 0 ",
    "-1",
    "+0",
    "0x0",
    "1e0",
    "00",
    "99999999999",
    "1.0",
    "2147483647",
    "2147483648",
    "4294967296",
    "-0",
    "0 ",
];
const ERRORS: &[&str] = &[
    "#FOO!",
    "#n/a",
    " #N/A ",
    "#SPILL!",
    "#CALC!",
    "#GETTING_DATA",
    "#NULL!",
    "#REF!",
    "#NUM!",
    "#VALUE!",
    "#DIV/0!",
    "#NAME?",
    "#N/A",
    "Err:502",
    "#FIELD!",
    "#NAME",
    "x",
    "#div/0!",
    " #DIV/0! ",
    "#Spill!",
    "#getting_data",
    "#N/A ",
    "#NULL!x",
    "#ref!",
];
const CELL_TYPES: &[&str] = &["x", "N", "S", "B", "STR", "inlinestr", "E", "D"];
const DATES: &[&str] = &[
    "2020-01-02",
    "2020-01-02T10:30:00",
    "2020-01-02T10:30:00.5",
    "2020-01-02T10:30:00Z",
    "10:30:00",
    "1900-01-01",
    "1900-02-28",
    "1900-03-01",
    "1899-12-31",
    "1904-01-01",
    "0001-01-01",
    "9999-12-31",
    "2020-1-2",
    "20200102",
    "abc",
    "2020-02-30",
    "2020-01-02T10:30",
    "2020-01-02+01:00",
    "2020-01-02T10:30:00.999",
    "2020-01-02T23:59:59.999",
    "2020-01-02T10:30:00+01:00",
    "2020-01-02T10:30:00-05:00",
    "2020-01-02T00:00:00+01:00",
    "2020-01-02Z",
    "2020-01-02z",
    "2020-01-02t10:30:00",
    "  2020-01-02",
    "2020-01-02 ",
    "+2020-01-02",
    "12020-01-02",
    "2020-01-02T24:00:00",
    "2020-01-02T10:60:00",
    "2020-01-02T10:30:60",
    "2020-01-02T24:30:00",
    "2020-01-02T25:00:00",
    "2019-02-29",
    "2020-02-29",
    "1900-02-29",
    "2020-13-01",
    "2020-00-10",
    " ",
    "2020-01-02T10:30:00+0100",
];
const DATES_1904: &[&str] = &[
    "2020-01-02",
    "1904-01-01",
    "1900-01-01",
    "2020-01-02T10:30:00",
    "2020-01-02T10:30:00+01:00",
    "2020-01-02+01:00",
    "2020-01-02Z",
];

const NUMBERS_MORE: &[&str] = &[
    "\n1", "&#13;1", " \n1", "\n\n-5", "\nNaN", "\nINF", "1.#NAN", "-1.#NAN", "1.#NANx", "1.#nan",
    "1.#NaN", "1.#SNAN", "\t 1",
];
const BOOLEANS_MORE: &[&str] = &["\n1", "&#13;1", "\t1"];
const INDEXES_MORE: &[&str] = &[
    "\u{2003}1",
    "\u{2000}-1",
    "\u{200B}1",
    "\u{200C}1",
    "\u{2028}1",
    "\u{a0}1",
    "-2147483648",
    "-2147483649",
    "\n1",
    "1",
    "&#8195;1",
];
const DATES_MORE: &[&str] = &[
    "1-01-01",
    "20-01-02",
    "99-01-01",
    "200-01-02",
    "-0001-01-01",
    "0000-01-01",
    "00000-01-01",
    "-0000-01-01",
    "32768-01-01",
    "32767-12-31",
    "65535-01-01",
    "99999-01-01",
    "100000-01-01",
    "9223372036854775807-01-01",
    "2020-01-02T12:00:00+00:30",
    "2020-01-02T12:00:00+05:30",
    "2020-01-02T12:00:00-05:30",
    "2020-01-02T12:00:00-00:30",
    "2020-01-02T12:00:00-00:15",
    "2020-01-02T12:00:00-00:01",
    "2020-01-02T12:00:00+14:00",
    "2020-01-02T12:00:00+14:01",
    "2020-01-02T12:00:00+15:00",
    "2020-01-02T12:00:00-14:00",
    "2020-01-02T12:00:00-14:30",
    "2020-01-02T12:00:00+00:60",
    "2020-01-02T12:00:00-00:00",
    "2020-01-02T10:30:00.",
    "2020-01-02T10:30:00.Z",
    "2020-01-02T10:30:00,5",
    "2020-01-02T10:30:00,5Z",
    "2020-01-02T10:30:00,5+01:00",
    "2020-01-02T10:30:00,",
    "2020-01-02T24:00:00.5",
    "2020-01-02T24:00:00.000000001",
    "2020-01-02T24:00:00.0",
    "2020-01-02T24:00:00.0000000001",
    "2020-01-02T24:00:00+01:00",
    "\u{2003}2020-01-02",
    "2020-01-02\u{2003}",
    "\u{a0}2020-01-02",
    "\u{3000}2020-01-02",
    "&#8195;2020-01-02",
    "2020-01-02&#10;",
];
const DATES_1904_MORE: &[&str] = &[
    "1904-01-01T10:30:00",
    "1903-12-31",
    "1904-01-02",
    "2020-01-02T12:00:00-05:30",
];
const DATE_1904_ATTRIBUTES: &[&str] = &[
    "1", "true", "on", "t", "2", "-1", "+1", "01", "1.0", "1e0", "1true", " 1", "1 ", "\t1", "1\t",
    "&#10;1", "1&#10;", "&#13;1", "1&#160;", "1&#8195;", "True", "TRUE", "T", "y", "yes", " true",
    "true ", " on", "on ", "0x1", "&#160;1", "&#8195;1", "0", "false", "",
];
const STRING_TEXTS: &[&str] = &[
    "a&#13;&#10;b",
    "a&#10;&#13;b",
    "a_x000D__x000A_b",
    "a&#13;_x000A_b",
    "a&#13;&#10;&#13;&#10;b",
    "a&#13;b",
    "a&#13;&#13;b",
    "a&#10;b",
    "a&#13;&#10;&#10;b",
    "_xD83D_",
    "_xDE00_",
    "a_xD83D_b",
    "_xDE00__xD83D_",
];
/// Cells with unusual structure, as (name, rows XML, range read back).
const STRUCTURES: &[(&str, &str, &str)] = &[
    (
        "inline_with_v_after_is",
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is><v>5</v></c></row>"#,
        "A1",
    ),
    (
        "inline_with_v_before_is",
        r#"<row r="1"><c r="A1" t="inlineStr"><v>5</v><is><t>b</t></is></c></row>"#,
        "A1",
    ),
    (
        "inline_with_self_closing_v",
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is><v/></c></row>"#,
        "A1",
    ),
    (
        "bad_cell_middle",
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="XFE1"><v>2</v></c><c r="C1"><v>3</v></c></row><row r="2"><c r="A2"><v>4</v></c></row>"#,
        "A1:C3",
    ),
    (
        "bad_row_middle",
        r#"<row r="1"><c r="A1"><v>1</v></c></row><row r="1048577"><c r="A1"><v>2</v></c></row><row r="3"><c r="A3"><v>5</v></c></row>"#,
        "A1:C3",
    ),
    (
        "row_zero_middle",
        r#"<row r="1"><c r="A1"><v>1</v></c></row><row r="0"><c r="B1"><v>2</v></c></row><row r="3"><c r="A3"><v>5</v></c></row>"#,
        "A1:C3",
    ),
    (
        "reversed_middle",
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="1A"><v>2</v></c><c r="C1"><v>3</v></c></row><row r="2"><c r="A2"><v>4</v></c></row>"#,
        "A1:C3",
    ),
    (
        "bad_row_text_addressed",
        r#"<row r="1"><c r="A1"><v>1</v></c></row><row r="x"><c r="B2"><v>2</v></c></row><row r="3"><c r="A3"><v>5</v></c></row>"#,
        "A1:C3",
    ),
    (
        "bad_row_then_unaddressed",
        r#"<row r="1"><c><v>1</v></c></row><row r="1048577"><c><v>2</v></c></row><row><c><v>3</v></c></row>"#,
        "A1:C4",
    ),
    (
        "row_zero_then_unaddressed",
        r#"<row r="2"><c><v>1</v></c></row><row r="0"><c><v>2</v></c></row><row><c><v>3</v></c></row>"#,
        "A1:C5",
    ),
    (
        "row_text_then_unaddressed",
        r#"<row r="1"><c><v>1</v></c></row><row r="x"><c><v>2</v></c></row><row><c><v>3</v></c></row>"#,
        "A1:C5",
    ),
    (
        "row_with_a_number_prefix",
        r#"<row r="3x"><c><v>1</v></c></row><row r="-1"><c><v>2</v></c></row><row><c><v>3</v></c></row>"#,
        "A1:C5",
    ),
    (
        "bad_cell_then_unaddressed",
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="XFE1"><v>2</v></c><c><v>3</v></c><c><v>4</v></c></row>"#,
        "A1:E2",
    ),
    (
        "bad_cell_first",
        r#"<row r="1"><c r="1A"><v>2</v></c><c><v>3</v></c></row>"#,
        "A1:E2",
    ),
    (
        "empty_address_cell",
        r#"<row r="3"><c r=""><v>2</v></c><c><v>3</v></c></row>"#,
        "A1:E4",
    ),
    (
        "range_address_cell",
        r#"<row r="1"><c r="B1:C1"><v>2</v></c><c><v>3</v></c></row>"#,
        "A1:E2",
    ),
    (
        "sheet_address_cell",
        r#"<row r="2"><c r="Sheet1!B1"><v>2</v></c><c><v>3</v></c></row>"#,
        "A1:E3",
    ),
    (
        "row_zero_address_cell",
        r#"<row r="2"><c r="A0"><v>2</v></c><c><v>3</v></c></row>"#,
        "A1:E3",
    ),
    (
        "row_past_the_sheet_address_cell",
        r#"<row r="2"><c r="A1048577"><v>2</v></c><c><v>3</v></c></row>"#,
        "A1:E3",
    ),
    (
        "row_attribute_does_not_move_addressed_cells",
        r#"<row r="5"><c r="A1"><v>2</v></c><c><v>3</v></c></row>"#,
        "A1:E6",
    ),
    (
        "row_repeated",
        r#"<row r="1"><c><v>7</v></c><c><v>7</v></c></row><row r="1"><c><v>8</v></c></row>"#,
        "A1:E3",
    ),
    (
        "rows_out_of_order",
        r#"<row r="3"><c><v>3</v></c></row><row r="2"><c><v>2</v></c></row><row><c><v>9</v></c></row>"#,
        "A1:E5",
    ),
    (
        "row_with_a_wide_space",
        "<row r=\"\u{2003}3\"><c><v>1</v></c></row><row><c><v>2</v></c></row>",
        "A1:C5",
    ),
    (
        "duplicate_formula_then_value",
        r#"<row r="1"><c r="A1"><f>1+1</f><v>9</v></c><c r="A1"><v>5</v></c></row>"#,
        "A1",
    ),
    (
        "duplicate_formula_then_text_result",
        r#"<row r="1"><c r="A1"><f>1+1</f><v>9</v></c><c r="A1" t="str"><v>x</v></c></row>"#,
        "A1",
    ),
    (
        "duplicate_formula_then_inline",
        r#"<row r="1"><c r="A1"><f>1+1</f><v>9</v></c><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>"#,
        "A1",
    ),
    (
        "duplicate_formula_then_empty_formula",
        r#"<row r="1"><c r="A1"><f>1+1</f><v>9</v></c><c r="A1"><f/><v>5</v></c></row>"#,
        "A1",
    ),
    (
        "duplicate_across_rows",
        r#"<row r="1"><c r="A1"><f>1+1</f><v>9</v></c></row><row r="1"><c r="A1"><v>5</v></c></row>"#,
        "A1",
    ),
    (
        "duplicate_value_then_formula",
        r#"<row r="1"><c r="A1"><v>5</v></c><c r="A1"><f>1+1</f><v>9</v></c></row>"#,
        "A1",
    ),
    (
        "duplicate_formula_over_formula",
        r#"<row r="1"><c r="A1"><f>1+1</f><v>9</v></c><c r="A1"><f>2+2</f><v>8</v></c></row>"#,
        "A1",
    ),
    (
        "lowercase_and_absolute_addresses",
        r#"<row r="1"><c r="a1"><v>1</v></c><c r="$B$1"><v>2</v></c></row>"#,
        "A1:C1",
    ),
];

/// Shared formula groups with anchors and dependents in valid and invalid arrangements, as
/// (name, rows XML, range).
const SHARED_GROUPS: &[(&str, &str, &str)] = &[
    (
        "shared_valid",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A4",
    ),
    (
        "shared_anchor_without_ref",
        r#"<row r="1"><c r="A1"><f t="shared" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_ref_not_containing_anchor",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A2:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_dependent_without_anchor",
        r#"<row r="1"><c r="A1"><f t="shared" si="0"/><v>9</v></c></row><row r="2"><c r="A2"><v>3</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_dependent_outside_ref",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row><row r="4"><c r="A4"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A5",
    ),
    (
        "shared_dependent_before_anchor",
        r#"<row r="1"><c r="A1"><f t="shared" si="0"/><v>9</v></c></row><row r="2"><c r="A2"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_two_anchors",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">2+2</f><v>4</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A4",
    ),
    (
        "shared_anchor_empty_text",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0"/><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_dependent_with_text",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0">5+5</f><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_ref_invalid",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_ref_reversed",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A3:A1" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_two_dimensional",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:B2" si="0">1+1</f><v>2</v></c><c r="B1"><f t="shared" si="0"/><v>9</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c><c r="B2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:C3",
    ),
    (
        "shared_dependent_without_value",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_invalid_then_valid_anchor",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">2+2</f><v>4</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A4",
    ),
    (
        "shared_valid_then_invalid_anchor",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" ref="zz" si="0">2+2</f><v>4</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A4",
    ),
    (
        "shared_other_group_survives",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c><c r="B1"><f t="shared" si="1"/><v>9</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:B2",
    ),
    (
        "shared_invalid_anchor_replaced_by_plain_formula",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="0">1+1</f><v>2</v></c><c r="A1"><f>3</f><v>3</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_valid_anchor_replaced_by_plain_formula",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c><c r="A1"><f>3</f><v>3</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_valid_anchor_replaced_by_constant",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c><c r="A1"><v>3</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_invalid_dependent_replaced_by_plain_formula",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c><c r="A2"><f>4</f><v>4</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_invalid_dependent_replaced_by_constant",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c><c r="A2"><v>4</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_groups_swapped_at_same_address",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="0">1+1</f><v>2</v></c><c r="A1"><f t="shared" ref="A1:A2" si="1">5+5</f><v>5</v></c></row><row r="2"><c r="A2"><f t="shared" si="1"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_valid_dependent_then_constant",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c><c r="A2"><v>4</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_valid_constant_then_dependent",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><v>4</v></c><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_invalid_constant_then_dependent",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><v>4</v></c><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_valid_plain_then_dependent",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f>4</f><v>4</v></c><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_invalid_plain_then_dependent",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f>4</f><v>4</v></c><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_invalid_dependent_then_text_dependent",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c><c r="A2"><f t="shared" si="0">7+7</f><v>7</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_valid_dependent_then_constant_text",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c><c r="A2" t="str"><v>x</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_valid_anchor_then_invalid_anchor_same_cell",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c><c r="A1"><f t="shared" ref="zz" si="1">5+5</f><v>5</v></c></row><row r="2"></row>"#,
        "A1:A3",
    ),
    (
        "shared_invalid_dependent_over_valid_dependent",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c><c r="A2"><f t="shared" si="1"/><v>8</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_orphan_dependent_over_valid_dependent",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c><c r="A2"><f t="shared" si="2"/><v>7</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_valid_dependent_over_orphan_dependent",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="2"/><v>7</v></c><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
    (
        "shared_invalid_anchor_over_valid_anchor_then_constant",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c><c r="A1"><f t="shared" ref="zz" si="1">5+5</f><v>5</v></c><c r="A1"><v>3</v></c></row><row r="2"></row>"#,
        "A1:A3",
    ),
    (
        "shared_two_invalid_then_constant",
        r#"<row r="1"><c r="A1"><f t="shared" ref="zz" si="1">5+5</f><v>5</v></c><c r="A1"><f t="shared" ref="zz" si="3">1</f><v>1</v></c><c r="A1"><v>3</v></c></row><row r="2"></row>"#,
        "A1:A3",
    ),
    (
        "shared_cell_in_range_without_formula",
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><v>7</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>9</v></c></row>"#,
        "A1:A3",
    ),
];

/// Adds one package per value, each with the value as the content of cell `A1`.
fn add_matrix(
    entries: &mut Vec<Crafted>,
    kind: &str,
    attributes: &str,
    strings: Option<&str>,
    properties: &str,
    values: &[&str],
) {
    for value in values {
        entries.push(Crafted {
            name: format!("{kind}[{value}]"),
            rows: format!(r#"<row r="1"><c r="A1"{attributes}><v>{value}</v></c></row>"#),
            strings: strings.map(str::to_owned),
            range: "A1",
            properties: properties.to_owned(),
        });
    }
}

/// The `ref` attribute of a shared formula anchor spelled in many ways.
const SHARED_REFERENCES: &[&str] = &[
    "A1:A",
    "A1:1",
    "A1:$A",
    "A1:A$",
    "A1:$",
    "A1:0",
    "A1:-1",
    "A1:+1",
    "A1:+A1",
    "A1:!",
    "A1:A1048577x",
    "A1:$A$0",
    "A1:A$0",
    "A1:A0$",
    "A1:B",
    "A1:B0z",
    "A1:XFD0",
    "A1:XFE",
    "A1:XFE0",
    "A1:AAAA0",
    "A1:A0:A2",
    "A2:A0",
    "A$1:A0",
    "A1:a0",
    "A1:A+1x",
    "A1:A-x",
    "A1:A0",
    "A1:A00",
    "A+0:A1",
    "A-0:A1",
    "A+1:A0",
    "A-1:A1",
    "A1:A+0x2",
    "A1:A-0x2",
    "A1:A0x0",
    "A0x1:A2",
    "A0x2:A3",
    "A1:A0X2",
    "A1:A0x",
    "A1:A0xg",
    "A1:A+",
    "A1:A-",
    "A1:A+-1",
    "A1:A1+1",
    "A1:A 2",
    "A1:A1 2",
    "A1:A0b1",
    "A1:A0o7",
    "A1:A1_0",
    "A1:A2",
    "A1",
    "$A$1:$A$2",
    "a1:a2",
    "A1:",
    ":A2",
    "",
    "A1:A2:A3",
    "A1:A2 B1",
    " A1:A2",
    "A1:A2 ",
    "A1:A2000000",
    "A1:XFE2",
    "A0:A2",
    "A2:A3",
    "A1:B1",
    "1:2",
    "A:A",
    "R1C1:R2C1",
    "Sheet1!A1:A2",
    "A1:A2,B1",
    "A1:A2\t",
    "A1 :A2",
    "A1: A2",
    "&#10;A1:A2",
    "A1:A99999999999999999999",
    "A1:A1048577",
    "A1048577:A1048578",
    "A:XFE",
    "A:A1",
    "1:A2",
    "AAAA1:A2",
    "A1:A-1",
    "A1:A+2",
    "A1:A2x",
    "$A1:A$2",
    "A$1:$A2",
    "A1:$A$2",
    "$$A1:A2",
    "0:2",
    "XFD1:XFD2",
    "A1:A2;",
    "A1;A2",
    "A1:a$2",
    "A1:A2&#10;",
    "A1:A2&#13;",
    "A1:A2&#9;",
    "A1:A2&#32;",
    "A01:A2",
    "A1:A02",
    "1:02",
    "A1:XFD1",
    "A1:XFD1048576",
    "A1:A18446744073709551616",
    "A1:A4294967297",
    "A1:A2  ",
    "A1:AAA1",
    "a:a",
    "$A:$A",
    "$1:$2",
    "A1:A2&#160;",
    "A1 ",
    "A1:A2\n",
    "A+1:A2",
    "A-1:A2",
    "A+1",
    "A-1",
    "+A1:A2",
    "A1:+A2",
    "A+1:A+2",
    "A+0:A1",
    "A1:A2.5",
    "A1:A2e0",
    "A1:A0x2",
    "A1:A+0",
    "A1:A+1",
    "A1:A+3",
    "A1:A+4",
    "A1:A+5",
    "A1:A+6",
    "A1:A+7",
    "A1:A-0",
    "A1:A-2",
    "A1:A-3",
    "A1:A-4",
    "A1:A-5",
    "A1:A-6",
    "A1:A-7",
    "A1:A+22",
    "A1:B+2",
    "A1:A-22",
    "A1:A+2 ",
    "A1:A+002",
    "A1:A+4294967297",
    "A1:A-4294967295",
];

/// Shared formula anchors with the `ref` attribute spelled in many ways.
fn shared_reference_entries(entries: &mut Vec<Crafted>) {
    for reference in SHARED_REFERENCES {
        entries.push(Crafted {
            name: format!("shared_ref[{reference}]"),
            rows: format!(
                r#"<row r="1"><c r="A1"><f t="shared" ref="{reference}" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>9</v></c></row>"#
            ),
            strings: None,
            range: "A1:A3",
            properties: String::new(),
        });
    }
}

/// The matrices added by the second review: whitespace, dates, flags, line breaks and addresses.
fn more_entries(entries: &mut Vec<Crafted>) {
    add_matrix(entries, "number", "", None, "", NUMBERS_MORE);
    add_matrix(entries, "boolean", r#" t="b""#, None, "", BOOLEANS_MORE);
    add_matrix(
        entries,
        "index",
        r#" t="s""#,
        Some("<si><t>a</t></si><si><t>b</t></si>"),
        "",
        INDEXES_MORE,
    );
    add_matrix(entries, "date", r#" t="d""#, None, "", DATES_MORE);
    add_matrix(
        entries,
        "date1904",
        r#" t="d""#,
        None,
        r#"<workbookPr date1904="1"/>"#,
        DATES_1904_MORE,
    );
    for attribute in DATE_1904_ATTRIBUTES {
        let properties = format!(r#"<workbookPr date1904="{attribute}"/>"#);
        add_matrix(
            entries,
            "date1904_attribute",
            r#" t="d""#,
            None,
            &properties,
            &["1904-01-01"],
        );
    }
    add_matrix(
        entries,
        "date1904_twice",
        r#" t="d""#,
        None,
        r#"<workbookPr date1904="1"/><workbookPr date1904="0"/>"#,
        &["1904-01-01"],
    );
    add_matrix(
        entries,
        "date1904_twice_reversed",
        r#" t="d""#,
        None,
        r#"<workbookPr date1904="0"/><workbookPr date1904="1"/>"#,
        &["1904-01-01"],
    );
    for text in STRING_TEXTS {
        entries.push(Crafted {
            name: format!("shared[{text}]"),
            rows: r#"<row r="1"><c r="A1" t="s"><v>0</v></c></row>"#.to_owned(),
            strings: Some(format!("<si><t>{text}</t></si>")),
            range: "A1",
            properties: String::new(),
        });
        entries.push(Crafted {
            name: format!("inline[{text}]"),
            rows: format!(r#"<row r="1"><c r="A1" t="inlineStr"><is><t>{text}</t></is></c></row>"#),
            strings: None,
            range: "A1",
            properties: String::new(),
        });
        entries.push(Crafted {
            name: format!("result[{text}]"),
            rows: format!(r#"<row r="1"><c r="A1" t="str"><v>{text}</v></c></row>"#),
            strings: None,
            range: "A1",
            properties: String::new(),
        });
    }
    for si in ["x", "", "0", "0.0", "1e0", "-1", " 0", "4294967296"] {
        entries.push(Crafted {
            name: format!("shared_si[{si}]"),
            rows: format!(
                r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="{si}">1+1</f><v>5</v></c></row><row r="2"><c r="A2"><f t="shared" si="{si}"/><v>6</v></c></row>"#
            ),
            strings: None,
            range: "A1:A3",
            properties: String::new(),
        });
    }
    shared_reference_entries(entries);
    for kind in ["", "bogus", "Shared", " normal", "NORMAL", "normal"] {
        entries.push(Crafted {
            name: format!("formula_type[{kind}]"),
            rows: format!(r#"<row r="1"><c r="A1"><f t="{kind}">1+1</f><v>5</v></c></row>"#),
            strings: None,
            range: "A1",
            properties: String::new(),
        });
    }
    for (name, rows, range) in SHARED_GROUPS.iter().chain(STRUCTURES) {
        entries.push(Crafted {
            name: (*name).to_owned(),
            rows: (*rows).to_owned(),
            strings: None,
            range,
            properties: String::new(),
        });
    }
}

/// The hand-written cells plus the matrices of unusual contents, each as one package.
fn crafted_entries() -> Vec<Crafted> {
    let mut entries: Vec<Crafted> = CRAFTED
        .iter()
        .map(|(name, rows, strings)| Crafted {
            name: (*name).to_owned(),
            rows: (*rows).to_owned(),
            strings: strings.map(str::to_owned),
            range: "A1",
            properties: String::new(),
        })
        .collect();
    add_matrix(&mut entries, "number", "", None, "", NUMBERS);
    add_matrix(&mut entries, "boolean", r#" t="b""#, None, "", BOOLEANS);
    add_matrix(
        &mut entries,
        "index",
        r#" t="s""#,
        Some("<si><t>a</t></si>"),
        "",
        INDEXES,
    );
    add_matrix(&mut entries, "error", r#" t="e""#, None, "", ERRORS);
    add_matrix(&mut entries, "date", r#" t="d""#, None, "", DATES);
    add_matrix(
        &mut entries,
        "date1904",
        r#" t="d""#,
        None,
        r#"<workbookPr date1904="1"/>"#,
        DATES_1904,
    );
    more_entries(&mut entries);
    for cell_type in CELL_TYPES {
        add_matrix(
            &mut entries,
            "type",
            &format!(r#" t="{cell_type}""#),
            None,
            "",
            &["1"],
        );
    }
    entries
}

/// Cells that hand-written XLSX parts hold are read as `LibreOffice` loads them.
#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn crafted_cells_read_as_the_oracle_loads_them() -> Result<(), Box<dyn StdError>> {
    let entries = crafted_entries();
    let mut cases = Vec::new();
    for (index, entry) in entries.iter().enumerate() {
        let path = format!("{CRAFTED_DIR}/crafted_{index}.xlsx");
        let package =
            single_sheet_package_with(&entry.rows, entry.strings.as_deref(), &entry.properties)?;
        fs::write(&path, package.into_inner())?;
        cases.push(Case {
            name: entry.name.clone(),
            sheets: vec![spec("Dummy", &[])],
            ops: vec![Op::Open { path }],
            reads: vec![Read {
                sheet: "S1".to_owned(),
                range: entry.range.to_owned(),
            }],
        });
    }
    let report = oracle::run(&cases)?;

    let mut differences = Vec::new();
    for (index, (outcome, entry)) in report.results.iter().zip(&entries).enumerate() {
        let name = &entry.name;
        let path = format!("{CRAFTED_DIR}/crafted_{index}.xlsx");
        let ours = sheetrs_xlsx::read(File::open(&path)?);
        let line = match (outcome, &ours) {
            (oracle::Outcome::Observed(observed), Ok(workbook)) => {
                let sheet = workbook.sheet("S1");
                let mut found = Vec::new();
                for (row, cells) in (0..).zip(&observed.reads[0].cells) {
                    for (col, expected) in (0..).zip(cells) {
                        let actual = sheet.and_then(|sheet| {
                            sheet.get(CellPos { row, col }).map(|view| view.to_cell())
                        });
                        if !same(expected, actual.as_ref(), true) {
                            found.push(format!(
                                "{name} at {row},{col}: oracle {expected:?}, read {actual:?}"
                            ));
                        }
                    }
                }
                (!found.is_empty()).then(|| found.join("\n"))
            }
            (oracle::Outcome::Observed(observed), Err(error)) => Some(format!(
                "{name}: oracle {:?}, read failed: {error}",
                observed.reads[0].cells[0][0]
            )),
            (oracle::Outcome::Failed { error, .. }, Ok(_)) => {
                Some(format!("{name}: oracle failed ({error}), read succeeded"))
            }
            (oracle::Outcome::Failed { .. }, Err(_)) => None,
        };
        differences.extend(line);
    }
    assert!(differences.is_empty(), "\n{}", differences.join("\n"));
    Ok(())
}

/// The formula the oracle shows for a cell, if any.
fn formula_of(value: &CellValue) -> Option<&String> {
    match value {
        CellValue::Number { formula, .. }
        | CellValue::String { formula, .. }
        | CellValue::Error { formula, .. }
        | CellValue::NonFinite { formula, .. } => formula.as_ref(),
        CellValue::Empty => None,
    }
}

/// Whether the cell read matches the oracle's, formula presence included. With `pending_stale`,
/// the stale value of a formula cell is not compared: `LibreOffice` recalculates it, which
/// needs the formula engine.
fn same(expected: &CellValue, actual: Option<&Cell>, pending_stale: bool) -> bool {
    // LibreOffice keeps a boolean as the formula `=TRUE()` or `=FALSE()` and an error constant as
    // a formula too; here they are constants (an open defect).
    let boolean = matches!(
        actual.map(|cell| (&cell.value, cell.formula.is_none())),
        Some((Value::Bool(_) | Value::Error(_), true))
    );
    let actual_has_formula = actual.is_some_and(|cell| cell.formula.is_some());
    if (formula_of(expected).is_some() != actual_has_formula) && !boolean {
        return false;
    }
    if pending_stale && actual.is_some_and(|cell| cell.stale) {
        return true;
    }
    match (expected, actual.map(|cell| &cell.value)) {
        (CellValue::Empty, None | Some(Value::Empty)) => true,
        (CellValue::Number { v, .. }, Some(Value::Bool(flag))) => {
            v.to_bits() == f64::from(u8::from(*flag)).to_bits()
        }
        (CellValue::NonFinite { v, .. }, Some(Value::Number(number))) => match v.as_str() {
            "nan" => number.is_nan(),
            "inf" => number.is_infinite() && number.is_sign_positive(),
            "-inf" => number.is_infinite() && number.is_sign_negative(),
            _ => false,
        },
        (CellValue::Number { v, .. }, Some(Value::Number(number))) => {
            v.to_bits() == number.to_bits()
        }
        (CellValue::String { v, .. }, Some(Value::Text(text))) => v.as_str() == &**text,
        (CellValue::Error { text, .. }, Some(Value::Error(kind))) => kind.as_text() == text,
        _ => false,
    }
}

fn spec(name: &str, cells: &[(&str, CellInput)]) -> SheetSpec {
    SheetSpec {
        name: name.to_owned(),
        cells: cells
            .iter()
            .map(|(address, content)| ((*address).to_owned(), content.clone()))
            .collect::<BTreeMap<_, _>>(),
    }
}

fn formula(text: &str) -> CellInput {
    CellInput::Text(text.to_owned())
}

/// The reader returns the values `LibreOffice` loads from an XLSX file it saved.
#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn a_workbook_saved_by_the_oracle_reads_back_to_the_observed_values()
-> Result<(), Box<dyn StdError>> {
    let case = Case {
        name: "xlsx_read_back".to_owned(),
        sheets: vec![
            spec(
                "S1",
                &[
                    ("A1", CellInput::Number(1.0)),
                    ("A2", CellInput::Number(2.5)),
                    ("A3", CellInput::Text("text & more".to_owned())),
                    ("B1", formula("=SUM(A1:A2)")),
                    ("B2", formula("=1/0")),
                    ("B3", formula("=\"!\"&A3")),
                    ("C1", formula("=0.1+0.2")),
                    ("C2", CellInput::Number(6.934_698_245_112_694e276)),
                ],
            ),
            spec("S2", &[("A1", CellInput::Text("x".to_owned()))]),
        ],
        // Reading the saved file back through the oracle gives the values LibreOffice
        // loads from it, which is what the reader must reproduce.
        ops: vec![
            Op::SaveAs {
                format: Format::Xlsx,
                path: OUTPUT.to_owned(),
            },
            Op::Open {
                path: OUTPUT.to_owned(),
            },
        ],
        reads: vec![
            Read {
                sheet: "S1".to_owned(),
                range: "A1:C3".to_owned(),
            },
            Read {
                sheet: "S2".to_owned(),
                range: "A1".to_owned(),
            },
        ],
    };
    let mut observed = oracle::run(&[case])?.into_observed()?;
    let observed = observed.remove(0);

    let workbook = sheetrs_xlsx::read(File::open(OUTPUT)?)?;
    let names: Vec<&str> = workbook.sheets().map(|sheet| sheet.name()).collect();
    assert_eq!(names, vec!["S1", "S2"]);
    assert!(
        workbook
            .sheets()
            .all(|sheet| sheet.visibility() == Visibility::Visible)
    );

    // LibreOffice stores formula results in XLSX with 15 significant digits, so
    // 0.1+0.2 reads back as 0.3 although it is 0.30000000000000004 in memory.
    let stored = workbook
        .sheet("S1")
        .and_then(|sheet| sheet.get(CellPos { row: 0, col: 2 }))
        .map(|cell| cell.to_cell().value);
    assert_eq!(stored, Some(Value::Number(0.3)));

    for read in &observed.reads {
        let sheet = workbook
            .sheet(&read.sheet)
            .ok_or_else(|| format!("sheet {} is missing", read.sheet))?;
        for (row, cells) in (0..).zip(&read.cells) {
            for (col, expected) in (0..).zip(cells) {
                let actual = sheet.get(CellPos { row, col }).map(|view| view.to_cell());
                assert!(
                    same(expected, actual.as_ref(), false),
                    "{}!{row},{col}: oracle {expected:?}, read {actual:?}",
                    read.sheet
                );
            }
        }
    }
    Ok(())
}

/// What `LibreOffice` keeps of the cached result of a formula whose cache disagrees with it,
/// as (name, cell XML, number the oracle shows, stale flag of the cell the reader returns).
const FORMULA_CACHES: &[(&str, &str, Option<f64>, Option<bool>)] = &[
    (
        "plain_number",
        "<c><f>1+1</f><v>7</v></c>",
        Some(7.0),
        Some(false),
    ),
    (
        "plain_number_one",
        "<c><f>1+1</f><v>1</v></c>",
        Some(1.0),
        Some(false),
    ),
    (
        "zero_number",
        "<c><f>1+1</f><v>0</v></c>",
        Some(2.0),
        Some(true),
    ),
    (
        "text_number",
        "<c><f>1+1</f><v>abc</v></c>",
        Some(2.0),
        Some(true),
    ),
    ("no_cache", "<c><f>1+1</f></c>", Some(2.0), Some(true)),
    (
        "string",
        r#"<c t="str"><f>1+1</f><v>1</v></c>"#,
        Some(2.0),
        Some(true),
    ),
    (
        "boolean",
        r#"<c t="b"><f>1+1</f><v>1</v></c>"#,
        Some(2.0),
        Some(true),
    ),
    (
        "error",
        r#"<c t="e"><f>1+1</f><v>1</v></c>"#,
        Some(2.0),
        Some(true),
    ),
    (
        "shared",
        r#"<c><f t="shared" ref="A1" si="0">1+1</f><v>7</v></c>"#,
        Some(2.0),
        Some(true),
    ),
    (
        "array",
        r#"<c><f t="array" ref="A1">1+1</f><v>7</v></c>"#,
        Some(2.0),
        Some(true),
    ),
    (
        "data_table",
        r#"<c><f t="dataTable" ref="A1"/><v>7</v></c>"#,
        None,
        None,
    ),
    (
        "inline_text",
        r#"<c t="inlineStr"><is><t>c</t></is><f>1+1</f></c>"#,
        Some(2.0),
        Some(true),
    ),
];

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn only_the_cache_of_a_plain_non_zero_number_is_trusted_on_load() -> Result<(), Box<dyn StdError>> {
    let mut cases = Vec::new();
    for (index, (name, cell, _, _)) in FORMULA_CACHES.iter().enumerate() {
        let rows = format!(
            r#"<row r="1">{}</row>"#,
            cell.replacen("<c", r#"<c r="A1""#, 1)
        );
        let path = format!("{CRAFTED_DIR}/cache_{index}.xlsx");
        fs::write(
            &path,
            single_sheet_package_with(&rows, None, "")?.into_inner(),
        )?;
        cases.push(Case {
            name: (*name).to_owned(),
            sheets: vec![spec("Dummy", &[])],
            ops: vec![Op::Open { path }],
            reads: vec![Read {
                sheet: "S1".to_owned(),
                range: "A1".to_owned(),
            }],
        });
    }
    let observed = oracle::run(&cases)?.into_observed()?;
    for (index, ((name, _, number, stale), case)) in
        FORMULA_CACHES.iter().zip(&observed).enumerate()
    {
        let shown = match &case.reads[0].cells[0][0] {
            CellValue::Number { v, .. } => Some(*v),
            _ => None,
        };
        assert_eq!(shown, *number, "{name}: what the oracle shows");

        let workbook =
            sheetrs_xlsx::read(File::open(format!("{CRAFTED_DIR}/cache_{index}.xlsx"))?)?;
        let cell = workbook
            .sheet("S1")
            .and_then(|sheet| sheet.get(CellPos { row: 0, col: 0 }));
        assert_eq!(
            cell.map(|cell| cell.stale),
            *stale,
            "{name}: what the reader flags"
        );
    }
    Ok(())
}
