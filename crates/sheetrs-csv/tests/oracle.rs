use std::error::Error as StdError;
use std::fs::File;

use sheetrs_core::Cell;
use sheetrs_csv::{FormulaFields, Options, QuotedFields, Spaces};
use sheetrs_testkit::case::{self, CellValue, TextFilter};
use sheetrs_testkit::compare::Leniency;
use sheetrs_testkit::crafted::{self, Package};

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
