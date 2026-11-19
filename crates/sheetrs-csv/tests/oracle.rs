use std::error::Error as StdError;
use std::fs::File;

use sheetrs_core::{Cell, CellPos, CellRef, ErrorKind, Formula, Value, Workbook};
use sheetrs_csv::{FormulaFields, Options, QuotedFields, Spaces, WriteOptions};
use sheetrs_testkit::case::{
    self, Case, CellInput, CellValue, Format, Op, Read, SheetSpec, TextFilter,
};
use sheetrs_testkit::compare::Leniency;
use sheetrs_testkit::crafted::{self, Package};
use sheetrs_testkit::oracle;

const DIR: &str = env!("CARGO_TARGET_TMPDIR");

/// Fields on their own, which the filter reads as numbers or as text.
const FIELDS: &[&str] = &[
    "1",
    "-1",
    "+1",
    "1.5",
    ".5",
    "1.",
    "1e3",
    "1E3",
    "1e+3",
    "007",
    "1 000",
    " 1",
    "1 ",
    "0x10",
    "NaN",
    "inf",
    "1e999",
    "-1e999",
    "#N/A",
    "TRUE",
    "true",
    "False",
    "=1+1",
    "-",
    "+",
    "'1",
    "1_000",
    "-0",
    "+0",
    "1e",
    "1e+",
    "--1",
    "+-1",
    "1.5.5",
    "abc",
    "00",
    "0.0",
    "1e-3",
    "-.5",
    "+.5e1",
    "٣",
    "１",
    "1d3",
    "1E3.5",
    "01/02/2020",
    "12:30",
    "50%",
    "$1.50",
    "1/2",
    "Jan 1",
    "1:00 PM",
    "(1)",
    "2020-1-2",
    "TRUE()",
    "#DIV/0!",
    "Err:502",
    "- 1",
    "1 e5",
    "1e 5",
    "+ 1",
    " +1",
    "1 1",
    "1\u{a0}",
    "\u{2003}1",
    "\u{3000}1",
    "1\u{a0}\u{a0}",
    "\u{a0}",
    "1\u{2009}",
    "\u{2009}1",
    "\u{200b}1",
    "1\t",
    "\t1",
    "\u{b}1",
    "\u{85}1",
    "\u{2028}1",
    "+\u{a0}1",
    "-\u{a0}1",
    "1e-999",
    "1e-400",
    "4.9e-324",
    "1.7976931348623157e308",
    "1.8e308",
    "123456789012345678901234567890",
    "0.30000000000000004",
    "9007199254740993",
    "0123",
    "00.50",
    "-007",
    "1e5",
    "1E-5",
    "1e+05",
    "1.e5",
    ".e5",
    "=",
    "é",
    "日本",
    "😀",
    "a\u{0}b",
    "1\u{0}",
    "   ",
    "\t",
];

const DATES: &[&str] = &[
    "2020-01-02",
    "2020-01-02T12:30:00",
    "2020-01-02T12:30",
    "2020-01-02Z",
    "20200102",
    "2020-13-01",
    "1900-02-29",
    "0001-01-01",
    "2020-01-02 ",
    " 2020-01-02",
    "-2020-01-02",
    "12020-01-02",
    "2020-01-02T12:30:00Z",
    "2020-01-02T12:30:00+01:00",
    "2020-01-02T12:30:00.5",
    "2020-01-02t12:30:00",
    "2020-02-30",
    "9999-12-31",
    "10000-01-01",
    "1899-12-30",
    "2020-00-01",
    "2020-1-02",
    "20-01-02",
    "0000-01-01",
    "-0001-01-01",
    "- 2020-01-02",
    "2020-01-02.",
    "1582-10-15",
    "1582-10-14",
    "1582-10-05",
    "1582-10-04",
    "1582-10-10",
    "1500-02-29",
    "1500-03-01",
    "1000-01-01",
    "0100-01-01",
    "0100-02-29",
    "0004-02-29",
    "-0004-02-29",
    "-0001-12-31",
    "-4712-01-01",
    "-32768-01-01",
    "-32767-01-01",
    "-20000-01-01",
    "-9999-01-01",
    "-10000-01-01",
    "32767-12-31",
    "32768-01-01",
    "99999-01-01",
    "1600-02-29",
    "2020-01-02T24:00:00",
    "2020-01-02T23:59:60",
    "2020-01-02T12:60:00",
    "2020-01-02T1:00:00",
    "2020-01-02T12:30:00.",
    "2020-01-02T12:30:00.123456789012",
    "2020-01-02\t",
    "\t2020-01-02",
    "2020-01-02\u{a0}",
    "\u{a0}2020-01-02",
    "2020-001-02",
    "02020-01-02",
    "2020-01-2",
    "2020-01-02T12:30:60",
    "2020-01-02T00:00:00",
    "2020-01-02T12:30:00 ",
    "2020-01-02  T12:30:00",
    "1582-10-15T12:00:00",
    "1000-01-01T12:00:00",
    "-0001-01-01T12:00:00",
    "0001-01-01T12:00:00",
    "2020-01-02T12:30:00.0",
    "2020-01-02T12:30:00,5",
    "-2020-1-1",
    "+2020-01-02",
    "2020-01-02T12:30:00.999999999",
    "2020-01-02T00:00:00.000000001",
    "-0002-03-01",
    "-0100-03-01",
    "0400-02-29",
    "0300-02-29",
    "1200-02-29",
    "1000-02-29",
    "1582-02-28",
    "1581-12-31",
];

/// Whole files: quotes, line breaks, rows of different lengths.
const TEXTS: &[(&str, &str)] = &[
    ("plain", "a,b,c\n1,2,3\n"),
    ("empty_fields", "a,,c\n,,\n"),
    ("trailing_separator", "a,b,\nc\n"),
    ("blank_lines", "\n\na\n"),
    ("crlf", "a\r\nb\r\n"),
    ("cr", "a\rb\r"),
    ("blank_crlf", "a\r\n\r\nb"),
    ("quoted_separator", "\"a,b\",c\n"),
    ("escaped_quote", "\"a\"\"b\",c\n"),
    ("quoted_newline", "\"a\nb\",c\nd\n"),
    ("text_after_quote", "\"a\"b,c\n"),
    ("quote_in_the_middle", "a\"b,c\n"),
    ("unterminated", "\"a,b\nc\n"),
    ("empty_quoted", "\"\",x\n"),
    ("empty_quoted_in_the_middle", "x,\"\",y\n"),
    ("space_before_quote", " \"a\" ,b\n"),
    (
        "quoted_numbers",
        "\"1\",\"007\",\" 1\",\"1.5\",\"2020-01-02\"\n",
    ),
    ("bom", "\u{feff}a,b\n"),
    ("double_bom", "\u{feff}\u{feff}a\n"),
    ("bom_and_quote", "\u{feff}\"a,b\",c\n"),
    ("semicolon", "a;b\n"),
    ("tab", "a\tb\n"),
    ("spaces", " a , 1 ,  b\n"),
    ("formulas", "=1+1,\"=2+2\", =3+3,=1+,=,\"=\"\n"),
    ("no_final_newline", "a,b\n1"),
    ("ragged", "a,b,c\n1\n1,2,3,4\n"),
    ("only_newline", "\n"),
    ("lone_quote", "\"\n"),
    ("quote_pair_then_text", "\"\"a\",b\n"),
    ("quote_then_space", "\"a\" ,b\n"),
    ("unterminated_then_closing_later", "\"a,b\nc\"d\ne\n"),
    ("unterminated_to_the_end", "\"a,b\nc,d\n"),
    ("doubled_quote_at_the_end", "\"a\"\""),
    ("triple_quote", "\"a\"\"\",b\n"),
    ("quote_inside_after_close", "\"a\"b\"c\",d\n"),
    ("second_field_quote_text", "x,\"b\"c\"\n"),
    ("quoted_then_separator_at_the_end", "\"a\",\n"),
    ("quoted_at_the_end", "\"a\""),
    ("quote_not_at_the_start", "x \"a\" y,z\n"),
    ("two_quoted_parts", "\"a\"  \"b\",c\n"),
    ("quoted_multiline_more_fields", "\"a\nb\",c,\"d\ne\"\nf,g\n"),
    ("tab_before_quote", "\t\"a\",b\n"),
    ("many_spaces_before_quote", "   \"a\",b\n"),
    ("space_then_text_after_close", "\"a\" b,c\n"),
    ("space_only_fields", "   ,\" \",x\n"),
    ("tab_only_field", "\t,x\n"),
    ("trailing_spaces_numbers", "1  ,  2,3\t\n"),
    ("quote_in_quoted_number", "\"1\"\"\",2\n"),
    ("cr_inside_quotes", "\"a\rb\",c\n"),
    ("crlf_inside_quotes", "\"a\r\nb\",c\n"),
    ("lf_cr_inside_quotes", "\"a\n\rb\",c\n"),
    ("lf_cr", "a\n\rb\n"),
    ("trailing_crlf_lf", "a\r\n\n"),
    ("only_spaces_line", "   \na\n"),
    ("many_columns", "a,b,c,d,e,f,g,h\n"),
    ("quote_after_separator_space", "a, \"b\"\n"),
    ("no_break_spaces", "a\u{a0}, \u{a0}\n"),
    ("blank_between_quoted", "\"a\"\n\n\"b\"\n"),
    ("quoted_empty_lines", "\"a\n\n\nb\"\n"),
];

fn filter(options: &Options) -> TextFilter {
    let mut filter = TextFilter::comma_separated();
    filter.separator = u32::from(options.separator);
    filter.delimiter = options.delimiter.map_or(0, u32::from);
    filter.quoted_fields = match options.quoted_fields {
        QuotedFields::Detected => case::QuotedFields::Detected,
        QuotedFields::Text => case::QuotedFields::Text,
    };
    filter.formula_fields = match options.formula_fields {
        FormulaFields::Text => case::FormulaFields::Text,
        FormulaFields::Evaluated => case::FormulaFields::Evaluated,
    };
    filter.spaces = match options.spaces {
        Spaces::Kept => case::Spaces::Kept,
        Spaces::Trimmed => case::Spaces::Trimmed,
    };
    filter
}

fn assert_reads_as_the_oracle(
    group: &str,
    options: &Options,
    files: &[(String, Vec<u8>, &str)],
) -> Result<(), Box<dyn StdError>> {
    let packages: Vec<Package> = files
        .iter()
        .enumerate()
        .map(|(index, (name, content, range))| Package {
            name: name.clone(),
            bytes: content.clone(),
            // The oracle names the sheet of a text file after the file.
            sheet: format!("{group}_{index}"),
            range: (*range).to_owned(),
        })
        .collect();
    let options = *options;
    crafted::assert_reads_as_the_oracle(
        DIR,
        group,
        "csv",
        Some(&filter(&options)),
        &packages,
        move |path| {
            let stem = path
                .rsplit('/')
                .next()
                .and_then(|file| file.strip_suffix(".csv"))
                .ok_or("a file name")?;
            Ok(sheetrs_csv::read(File::open(path)?, stem, &options)?)
        },
        |_: &CellValue, _: Option<&Cell>| Leniency::Strict,
    )
}

fn field_files(fields: &[&str]) -> Vec<(String, Vec<u8>, &'static str)> {
    fields
        .iter()
        .map(|field| {
            (
                format!("field[{field}]"),
                format!("{field}\n").into_bytes(),
                "A1:A1",
            )
        })
        .collect()
}

fn text_files() -> Vec<(String, Vec<u8>, &'static str)> {
    TEXTS
        .iter()
        .map(|(name, content)| ((*name).to_owned(), content.as_bytes().to_vec(), "A1:F5"))
        .collect()
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn plain_numbers_and_text_over_csv() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("fields", &Options::comma_separated(), &field_files(FIELDS))
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn iso_dates_over_csv() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("dates", &Options::comma_separated(), &field_files(DATES))
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn quotes_line_breaks_and_rows_over_csv() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("texts", &Options::comma_separated(), &text_files())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn quoted_fields_as_text_over_csv() -> Result<(), Box<dyn StdError>> {
    let options = Options {
        quoted_fields: QuotedFields::Text,
        ..Options::comma_separated()
    };
    assert_reads_as_the_oracle("quoted_text", &options, &text_files())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn trimmed_spaces_over_csv() -> Result<(), Box<dyn StdError>> {
    let options = Options {
        spaces: Spaces::Trimmed,
        ..Options::comma_separated()
    };
    let mut files = text_files();
    files.extend(field_files(FIELDS));
    assert_reads_as_the_oracle("trimmed", &options, &files)
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn formula_fields_over_csv() -> Result<(), Box<dyn StdError>> {
    let options = Options {
        formula_fields: FormulaFields::Evaluated,
        ..Options::comma_separated()
    };
    let mut files = text_files();
    files.extend(field_files(&[
        "=1+1",
        "=1+",
        "=",
        "=A1",
        "=SUM(1,2)",
        "= 1",
        " =1",
        "=\"a\"",
        "=1/0",
    ]));
    assert_reads_as_the_oracle("formulas", &options, &files)
}

/// The same files with another separator and delimiter written in place of the comma and the
/// double quote.
#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn other_separators_and_delimiters_over_csv() -> Result<(), Box<dyn StdError>> {
    for (group, separator, delimiter) in [
        ("semicolon", ';', Some('"')),
        ("tab", '\t', Some('"')),
        ("pipe_and_apostrophe", '|', Some('\'')),
        ("no_delimiter", ',', None),
        ("space", ' ', Some('"')),
    ] {
        let options = Options {
            separator,
            delimiter,
            ..Options::comma_separated()
        };
        let files: Vec<_> = text_files()
            .into_iter()
            .map(|(name, content, range)| {
                let content = String::from_utf8_lossy(&content).into_owned();
                let mut written = content
                    .replace(',', "\u{1}")
                    .replace('"', &delimiter.unwrap_or('"').to_string());
                written = written
                    .replace(separator, "\u{2}")
                    .replace('\u{1}', &separator.to_string());
                (name, written.replace('\u{2}', ",").into_bytes(), range)
            })
            .collect();
        assert_reads_as_the_oracle(group, &options, &files)?;
    }
    Ok(())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn bytes_that_are_not_utf8_over_csv() -> Result<(), Box<dyn StdError>> {
    let files = vec![
        ("lone_byte".to_owned(), b"x\xff\xfey,a\n".to_vec(), "A1:C1"),
        (
            "bad_continuation".to_owned(),
            b"\xc3(,\xe2\x82\n".to_vec(),
            "A1:C1",
        ),
        (
            "truncated_sequence".to_owned(),
            b"ab\xf0\x9f".to_vec(),
            "A1:C1",
        ),
    ];
    assert_reads_as_the_oracle("bytes", &Options::comma_separated(), &files)
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn the_edge_of_the_sheet_over_csv() -> Result<(), Box<dyn StdError>> {
    let wide: Vec<String> = (0..16_390).map(|number| number.to_string()).collect();
    let tall: Vec<String> = (0..1_048_580).map(|number| number.to_string()).collect();
    let files = vec![
        (
            "columns".to_owned(),
            wide.join(",").into_bytes(),
            "XFB1:XFD1",
        ),
        (
            "rows".to_owned(),
            tall.join("\n").into_bytes(),
            "A1048575:A1048576",
        ),
    ];
    assert_reads_as_the_oracle("edge", &Options::comma_separated(), &files)
}

/// A cell of an exported sheet: what the oracle is given, and the value it shows as.
struct Exported {
    address: String,
    input: CellInput,
    value: Value,
}

fn constant(address: &str, input: CellInput, value: Value) -> Exported {
    Exported {
        address: address.to_owned(),
        input,
        value,
    }
}

fn number_cell(address: &str, number: f64) -> Exported {
    constant(address, CellInput::Number(number), Value::Number(number))
}

fn text_cell(address: &str, text: &str) -> Exported {
    constant(
        address,
        CellInput::Text(text.to_owned()),
        Value::Text(text.into()),
    )
}

fn formula_cell(address: &str, formula: &str, value: Value) -> Exported {
    constant(address, CellInput::Text(formula.to_owned()), value)
}

/// Numbers of every magnitude and number of digits, from a fixed sequence.
fn spread_numbers() -> Vec<f64> {
    let mut state: u64 = 0x2545_F491_4F6C_DD1D;
    let mut next = move || {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        state
    };
    let mut numbers = Vec::new();
    for digits in 1_usize..=17 {
        for exponent in -16_i32..=18 {
            let mut mantissa = String::new();
            for place in 0..digits {
                // The first digit is not zero.
                let digit = u8::try_from(next() % 10).unwrap_or(0);
                mantissa.push(char::from(
                    b'0' + if place == 0 { digit.max(1) } else { digit },
                ));
            }
            let text = format!("{}.{}e{exponent}", &mantissa[..1], &mantissa[1..]);
            if let Ok(number) = text.parse::<f64>() {
                numbers.push(if next() % 5 == 0 { -number } else { number });
            }
        }
    }
    numbers.extend([
        0.0,
        5e-324,
        f64::MAX,
        1e15,
        1e16,
        1e-4,
        1e-9,
        1e-10,
        999_999_999_999_999.9,
    ]);
    numbers
}

fn number_sheet(numbers: &[f64]) -> Vec<Exported> {
    numbers
        .iter()
        .enumerate()
        .map(|(row, number)| number_cell(&format!("A{}", row + 1), *number))
        .collect()
}

fn text_sheet() -> Vec<Exported> {
    let texts = [
        "plain",
        "with,comma",
        "with \"quote\"",
        "line\nbreak",
        " lead",
        "trail ",
        "123",
        "TRUE",
        "a;b",
        "tab\there",
        "\'apos",
        "\u{e9} \u{f1} \u{1f600}",
        "  ",
        "\"",
        "-",
        "1,5",
        "#N/A",
    ];
    texts
        .iter()
        .enumerate()
        .map(|(row, text)| text_cell(&format!("A{}", row + 1), text))
        .collect()
}

fn shape_sheet() -> Vec<Exported> {
    vec![
        number_cell("B2", 1.0),
        text_cell("D2", "x"),
        formula_cell("A4", "=TRUE()", Value::Bool(true)),
        formula_cell("B4", "=FALSE()", Value::Bool(false)),
        formula_cell("C4", "=1/0", Value::Error(ErrorKind::DivZero)),
        formula_cell("D4", "=NA()", Value::Error(ErrorKind::Na)),
        formula_cell("E4", "=\"\"", Value::Text("".into())),
        formula_cell(
            "A6",
            "=SQRT(-1)",
            Value::Error(ErrorKind::Other("Err:502".into())),
        ),
        number_cell("F7", 2.0),
    ]
}

/// The separator and delimiter of an export, as the oracle is given them and as the writer is.
const FORMATS: &[(&str, &str, WriteOptions)] = &[
    (
        "comma",
        "44,34,76,1,,1033,false,true,true,false,false,0",
        WriteOptions {
            separator: ',',
            delimiter: '"',
            quote_all_text: false,
            byte_order_mark: false,
        },
    ),
    (
        "semicolon",
        "59,34,76,1,,1033,false,true,true,false,false,0",
        WriteOptions {
            separator: ';',
            delimiter: '"',
            quote_all_text: false,
            byte_order_mark: false,
        },
    ),
    (
        "tab",
        "9,39,76,1,,1033,false,true,true,false,false,0",
        WriteOptions {
            separator: '\t',
            delimiter: '\'',
            quote_all_text: false,
            byte_order_mark: false,
        },
    ),
    (
        "quoted",
        "44,34,76,1,,1033,true,true,true,false,false,0",
        WriteOptions {
            separator: ',',
            delimiter: '"',
            quote_all_text: true,
            byte_order_mark: false,
        },
    ),
    (
        "bom",
        "44,34,76,1,,1033,false,true,true,false,false,0,false,true",
        WriteOptions {
            separator: ',',
            delimiter: '"',
            quote_all_text: false,
            byte_order_mark: true,
        },
    ),
];

/// The text export writes what `LibreOffice` writes: byte for byte.
#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn exported_text_is_what_the_oracle_exports_over_csv() -> Result<(), Box<dyn StdError>> {
    let sheets: Vec<(&str, Vec<Exported>)> = vec![
        ("numbers", number_sheet(&spread_numbers())),
        ("texts", text_sheet()),
        ("shape", shape_sheet()),
    ];
    let mut cases = Vec::new();
    let mut expected = Vec::new();
    for (sheet_name, cells) in &sheets {
        for (format_name, options, write_options) in FORMATS {
            let name = format!("{sheet_name}_{format_name}");
            let path = format!("{DIR}/export_{name}.csv");
            cases.push(Case {
                name: name.clone(),
                sheets: vec![SheetSpec {
                    name: "S1".to_owned(),
                    cells: cells
                        .iter()
                        .map(|cell| (cell.address.clone(), cell.input.clone()))
                        .collect(),
                }],
                ops: vec![Op::SaveAs {
                    format: Format::Csv,
                    path: path.clone(),
                    filter_options: Some((*options).to_owned()),
                }],
                reads: vec![Read {
                    sheet: "S1".to_owned(),
                    range: "A1:A1".to_owned(),
                }],
            });
            let mut workbook = Workbook::new();
            let mut sheet = workbook.sheet_builder("S1");
            for cell in cells {
                let at = CellRef::from_a1(&cell.address)?;
                let formula = match &cell.input {
                    CellInput::Text(text) if text.starts_with('=') => {
                        Some(Formula::Text(text.clone()))
                    }
                    _ => None,
                };
                sheet.set(
                    CellPos {
                        row: at.row.index,
                        col: at.col.index,
                    },
                    &Cell {
                        value: cell.value.clone(),
                        formula,
                        stale: false,
                    },
                )?;
            }
            sheet.finish();
            let mut written = Vec::new();
            let view = workbook.sheet("S1").ok_or("sheet")?;
            sheetrs_csv::write(view, write_options, &mut written)?;
            expected.push((name, path, written));
        }
    }
    oracle::run(&cases)?.into_observed()?;
    let mut differences = Vec::new();
    for (name, path, ours) in expected {
        let theirs = std::fs::read(&path)?;
        if theirs != ours {
            let (a, b) = (
                String::from_utf8_lossy(&theirs),
                String::from_utf8_lossy(&ours),
            );
            let line = a.lines().zip(b.lines()).position(|(x, y)| x != y);
            let input = line.and_then(|at| spread_numbers().get(at).copied());
            differences.push(format!(
                "{name}: input {input:?} first differing line {line:?}: oracle {:?}, written {:?}",
                line.and_then(|at| a.lines().nth(at)),
                line.and_then(|at| b.lines().nth(at))
            ));
        }
    }
    assert!(differences.is_empty(), "\n{}", differences.join("\n"));
    Ok(())
}
