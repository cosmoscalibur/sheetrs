use std::error::Error as StdError;
use std::fs::File;

use sheetrs_core::Cell;
use sheetrs_testkit::case::CellValue;
use sheetrs_testkit::compare::Leniency;
use sheetrs_testkit::crafted::{self, Package};

mod common;
use common::{content_xml, document as document_xml, package};

const DIR: &str = env!("CARGO_TARGET_TMPDIR");

/// A hand-written table: the rows XML of sheet `S1` and the range to compare.
#[derive(Clone)]
struct Crafted {
    name: String,
    rows: String,
    settings: String,
    range: &'static str,
    /// Whether the document names its namespaces with other prefixes.
    renamed: bool,
    /// A whole `content.xml`, which replaces the one built from the rows.
    raw: Option<String>,
}

fn entry(name: impl Into<String>, rows: impl Into<String>, range: &'static str) -> Crafted {
    Crafted {
        name: name.into(),
        rows: rows.into(),
        settings: String::new(),
        range,
        renamed: false,
        raw: None,
    }
}

/// One row of one cell, as the cell alone.
fn single(name: impl Into<String>, cell: &str) -> Crafted {
    entry(
        name,
        format!("<table:table-row>{cell}</table:table-row>"),
        "A1:D1",
    )
}

fn float(value: &str) -> String {
    format!(r#"<table:table-cell office:value-type="float" office:value="{value}"/>"#)
}

/// The `content.xml` of `entry`, with the prefixes it asks for: only the names of elements and
/// attributes change, not the namespaces they are in.
fn document(entry: &Crafted) -> String {
    if let Some(raw) = &entry.raw {
        return raw.clone();
    }
    let mut content = content_xml(&entry.rows, &entry.settings);
    if entry.renamed {
        for (prefix, other) in [("table", "tbl"), ("office", "ofc"), ("text", "txt")] {
            for lead in ["<", "</", " "] {
                content = content.replace(&format!("{lead}{prefix}:"), &format!("{lead}{other}:"));
            }
            content = content.replace(&format!("xmlns:{prefix}="), &format!("xmlns:{other}="));
        }
    }
    content
}

fn assert_reads_as_the_oracle(group: &str, entries: &[Crafted]) -> Result<(), Box<dyn StdError>> {
    let mut packages = Vec::new();
    for entry in entries {
        packages.push(Package {
            name: entry.name.clone(),
            bytes: package(&document(entry))?.into_inner(),
            sheet: "S1".to_owned(),
            range: entry.range.to_owned(),
        });
    }
    crafted::assert_reads_as_the_oracle(
        DIR,
        group,
        "ods",
        None,
        &packages,
        |path| Ok(sheetrs_ods::read(File::open(path)?)?),
        |_: &CellValue, _: Option<&Cell>| Leniency::Strict,
    )
}

const FLOATS: &[&str] = &[
    "abc",
    " 1 ",
    "1e3",
    "0x1",
    "inf",
    "NaN",
    "",
    "+1",
    ".5",
    "1,5",
    "1.",
    "-0",
    "-",
    "1e999",
    "-1e999",
    "١",
    "1 2",
    "0.1e1x",
    "\t2",
    "\n3",
    "\r4",
    "0.1",
    "1e-999",
    "1E5",
    "1e+2",
    "9007199254740993",
    "0.30000000000000004",
    "123456789012345678901234567890",
    ".",
    "+.5",
    "-.5e1",
    "5e",
    "5e+",
];

const BOOLEANS: &[&str] = &[
    "true", "TRUE", "True", "false", "0", "1", "yes", "", " true", "2", "-1", "t", "1.5", "0.0",
    "-0", "true ", "inf", "1e0",
];

const DATES: &[&str] = &[
    "2020-01-02",
    "2020-01-02Z",
    "2020-01-02+01:00",
    "2020-1-2",
    " 2020-01-02",
    "2020-01-02T25:00:00",
    "-0001-01-01",
    "12020-01-02",
    "2020-01-02T12:00",
    "2020-01-02T12:00:00.5",
    "",
    "2020-13-01",
    "2020-02-30",
    "abc",
    "2020-01-02T12:00:00Z",
    "2020-01-02 12:00:00",
    "1899-12-30",
    "1900-01-01",
    "1900-02-28",
    "1900-03-01",
    "0001-01-01",
    "9999-12-31",
    "2020-01-02T12:00:00+05:30",
    "2020-01-02T12:00:00-05:30",
    "2020-01-02T00:00:00.123456789",
    "2020-01-02T23:59:59.999",
    "2020-01-02T24:00:00",
    "2020-01-02T12:60:00",
    "2020-01-02T12:00:60",
    "2020-01-02T12:00:00+14:00",
    "2020-01-02T12:00:00+15:00",
    "2020-01-02+24:00",
    "0000-01-01",
    "-0000-01-01",
    "+2020-01-02",
    "2020-01-02 ",
    "2020-01-02T",
    "2020-01-02t12:00:00",
    "32768-01-01",
    "-32768-01-01",
    "2020-00-01",
    "2020-01-00",
    "2020-1-02",
    "20-01-02",
    "2020-01-02T1:00:00",
    "1582-10-10",
    "1582-10-04",
    "1900-02-29",
    "4712-01-01",
    "-0004-12-31",
    "-0005-01-01",
    "-0400-12-31",
    "-2020-01-02",
    "-4712-01-01",
    "-0100-03-01",
    "-0002-03-01",
    "2020-01-02T12:00:00.0000000001",
    "2020-01-02T12:00:00,5",
    "2020-01-02T12:00:00.",
];

const TIMES: &[&str] = &[
    "PT1H",
    "PT1H30M",
    "P1DT1H",
    "PT-1H",
    "-PT1H",
    "PT1.5S",
    "PT36H",
    "1:00",
    "",
    "PT0S",
    "P1D",
    "PT1M1S",
    "PT1H1H",
    "pt1h",
    "PT 1H",
    "PT0.0001S",
    "P1Y",
    "P1M",
    "P1Y1M",
    "PT1H1M1S",
    "P0D",
    "PT",
    "P",
    "PT1S",
    "PT0.5S",
    "P1DT0H",
    "PT60M",
    "PT1M60S",
    "PT24H",
    "PT1H.5S",
    "PT.5S",
    "PT1,5S",
    "-P1D",
    "+PT1H",
    "PT-0S",
    "PT1H-1M",
    "PT1.S",
    "PT100000H",
    "P100000D",
    "PT1H2H",
    "PT2H1H",
    "PT1S1M",
    "P1DT",
    "PT1.5H",
    "PT1.5M",
    "P1.5D",
    "-PT0S",
    "PT0.00000001S",
    "p1dt1h",
    "PT1h1m1s",
    "PT99999999999H",
];

const TYPES: &[&str] = &[
    "void",
    "foo",
    "FLOAT",
    " float",
    "currency",
    "percentage",
    "float",
    "string",
    "boolean",
    "date",
    "time",
    "",
];

fn scalar_entries() -> Vec<Crafted> {
    let mut entries = Vec::new();
    for value in FLOATS {
        entries.push(single(
            format!("float[{value}]"),
            &format!(r#"<table:table-cell office:value-type="float" office:value="{value}"/>"#),
        ));
    }
    entries.push(single(
        "float_no_value",
        r#"<table:table-cell office:value-type="float"/>"#,
    ));
    entries.push(single(
        "float_text_only",
        r#"<table:table-cell office:value-type="float"><text:p>7</text:p></table:table-cell>"#,
    ));
    entries.push(single(
        "percentage",
        r#"<table:table-cell office:value-type="percentage" office:value="0.25"/>"#,
    ));
    entries.push(single(
        "currency",
        r#"<table:table-cell office:value-type="currency" office:value="3" office:currency="USD"/>"#,
    ));
    for value in BOOLEANS {
        entries.push(single(
            format!("boolean[{value}]"),
            &format!(
                r#"<table:table-cell office:value-type="boolean" office:boolean-value="{value}"/>"#
            ),
        ));
    }
    entries.push(single(
        "boolean_none",
        r#"<table:table-cell office:value-type="boolean"/>"#,
    ));
    entries
}

/// Date and time cells.
fn date_entries() -> Vec<Crafted> {
    let mut entries = Vec::new();
    for value in DATES {
        entries.push(single(
            format!("date[{value}]"),
            &format!(r#"<table:table-cell office:value-type="date" office:date-value="{value}"/>"#),
        ));
    }
    entries.push(single(
        "date_with_value",
        r#"<table:table-cell office:value-type="date" office:value="3"/>"#,
    ));
    for value in TIMES {
        entries.push(single(
            format!("time[{value}]"),
            &format!(r#"<table:table-cell office:value-type="time" office:time-value="{value}"/>"#),
        ));
    }
    entries
}

/// Cells whose type attributes disagree or are not types.
fn typing_entries() -> Vec<Crafted> {
    let mut entries = Vec::new();
    for kind in TYPES {
        entries.push(single(
            format!("type[{kind}]"),
            &format!(
                r#"<table:table-cell office:value-type="{kind}" office:value="1"><text:p>t</text:p></table:table-cell>"#
            ),
        ));
    }
    for (name, cell) in [
        (
            "calcext_string_over_float",
            r#"<table:table-cell office:value-type="float" office:value="3" calcext:value-type="string"><text:p>t</text:p></table:table-cell>"#,
        ),
        (
            "calcext_float_over_string",
            r#"<table:table-cell office:value-type="string" calcext:value-type="float" office:value="4"><text:p>t</text:p></table:table-cell>"#,
        ),
        (
            "calcext_error_over_float",
            r#"<table:table-cell office:value-type="float" office:value="3" calcext:value-type="error"/>"#,
        ),
        (
            "calcext_error_string",
            r#"<table:table-cell office:value-type="string" calcext:value-type="error"><text:p>#DIV/0!</text:p></table:table-cell>"#,
        ),
        (
            "calcext_boolean",
            r#"<table:table-cell office:value-type="string" calcext:value-type="boolean"><text:p>t</text:p></table:table-cell>"#,
        ),
        (
            "calcext_same",
            r#"<table:table-cell office:value-type="float" office:value="3" calcext:value-type="float"/>"#,
        ),
        (
            "calcext_unknown",
            r#"<table:table-cell office:value-type="float" office:value="3" calcext:value-type="foo"/>"#,
        ),
        (
            "calcext_empty",
            r#"<table:table-cell office:value-type="float" office:value="3" calcext:value-type=""/>"#,
        ),
        (
            "value_on_string",
            r#"<table:table-cell office:value-type="string" office:value="3"><text:p>t</text:p></table:table-cell>"#,
        ),
        (
            "string_value_on_float",
            r#"<table:table-cell office:value-type="float" office:string-value="s" office:value="3"/>"#,
        ),
    ] {
        entries.push(single(name, cell));
    }
    entries
}

const STRINGS: &[(&str, &str)] = &[
    (
        "attribute_over_paragraph",
        r#"<table:table-cell office:value-type="string" office:string-value="a"><text:p>b</text:p></table:table-cell>"#,
    ),
    (
        "two_paragraphs",
        r#"<table:table-cell office:value-type="string"><text:p>hi</text:p><text:p>there</text:p></table:table-cell>"#,
    ),
    (
        "spaces_tabs_breaks",
        r#"<table:table-cell office:value-type="string"><text:p>a<text:s text:c="3"/>b<text:tab/>c<text:line-break/>d</text:p></table:table-cell>"#,
    ),
    (
        "no_content",
        r#"<table:table-cell office:value-type="string"/>"#,
    ),
    (
        "empty_attribute_uses_paragraph",
        r#"<table:table-cell office:value-type="string" office:string-value=""><text:p>x</text:p></table:table-cell>"#,
    ),
    (
        "empty_paragraph",
        r#"<table:table-cell office:value-type="string"><text:p/></table:table-cell>"#,
    ),
    (
        "two_empty_paragraphs",
        r#"<table:table-cell office:value-type="string"><text:p/><text:p/></table:table-cell>"#,
    ),
    (
        "padded",
        r#"<table:table-cell office:value-type="string"><text:p>  a  </text:p></table:table-cell>"#,
    ),
    (
        "span_and_link",
        r#"<table:table-cell office:value-type="string"><text:p>a<text:span>b</text:span><text:a xlink:href="x">c</text:a></text:p></table:table-cell>"#,
    ),
    (
        "heading_ignored",
        r#"<table:table-cell office:value-type="string"><text:h>head</text:h><text:p>x</text:p></table:table-cell>"#,
    ),
    (
        "list_ignored",
        r#"<table:table-cell office:value-type="string"><text:list><text:list-item><text:p>i</text:p></text:list-item></text:list></table:table-cell>"#,
    ),
    (
        "space_default_count",
        r#"<table:table-cell office:value-type="string"><text:p>a<text:s/>b</text:p></table:table-cell>"#,
    ),
    (
        "space_zero_count",
        r#"<table:table-cell office:value-type="string"><text:p>a<text:s text:c="0"/>b</text:p></table:table-cell>"#,
    ),
    (
        "space_bad_count",
        r#"<table:table-cell office:value-type="string"><text:p>a<text:s text:c="x"/>b</text:p></table:table-cell>"#,
    ),
    (
        "space_negative_count",
        r#"<table:table-cell office:value-type="string"><text:p>a<text:s text:c="-2"/>b</text:p></table:table-cell>"#,
    ),
    (
        "space_prefix_count",
        r#"<table:table-cell office:value-type="string"><text:p>a<text:s text:c="2x"/>b</text:p></table:table-cell>"#,
    ),
    (
        "carriage_return_line_feed",
        "<table:table-cell office:value-type=\"string\"><text:p>a\r\nb</text:p></table:table-cell>",
    ),
    (
        "lone_carriage_return",
        "<table:table-cell office:value-type=\"string\"><text:p>a\rb</text:p></table:table-cell>",
    ),
    (
        "line_feed_carriage_return",
        "<table:table-cell office:value-type=\"string\"><text:p>a\n\rb</text:p></table:table-cell>",
    ),
    (
        "literal_line_feed_and_tab",
        "<table:table-cell office:value-type=\"string\"><text:p>a\nb\tc</text:p></table:table-cell>",
    ),
    (
        "entities",
        r#"<table:table-cell office:value-type="string"><text:p>a&amp;b&lt;&#65;</text:p></table:table-cell>"#,
    ),
    (
        "annotation_ignored",
        r#"<table:table-cell office:value-type="string"><office:annotation><text:p>note</text:p></office:annotation><text:p>x</text:p></table:table-cell>"#,
    ),
    (
        "space_run",
        r#"<table:table-cell office:value-type="string"><text:p>a     b</text:p></table:table-cell>"#,
    ),
    (
        "without_type",
        r"<table:table-cell><text:p>abc</text:p></table:table-cell>",
    ),
    (
        "without_type_empty_paragraph",
        r"<table:table-cell><text:p/></table:table-cell>",
    ),
    (
        "without_type_blank_paragraph",
        r"<table:table-cell><text:p> </text:p></table:table-cell>",
    ),
    (
        "nested_paragraph_markup",
        r#"<table:table-cell office:value-type="string"><text:p><text:span><text:span>deep</text:span></text:span></text:p></table:table-cell>"#,
    ),
    (
        "note_in_paragraph",
        r#"<table:table-cell office:value-type="string"><text:p>a<text:note><text:note-body><text:p>n</text:p></text:note-body></text:note>b</text:p></table:table-cell>"#,
    ),
    (
        "cdata",
        r#"<table:table-cell office:value-type="string"><text:p><![CDATA[x<y]]></text:p></table:table-cell>"#,
    ),
    (
        "non_breaking_space",
        "<table:table-cell office:value-type=\"string\"><text:p>a\u{a0}b</text:p></table:table-cell>",
    ),
    (
        "paragraph_after_other_element",
        r#"<table:table-cell office:value-type="string"><text:soft-page-break/><text:p>x</text:p></table:table-cell>"#,
    ),
    (
        "nested_span",
        r#"<table:table-cell office:value-type="string"><text:p><text:span>b<text:span>c</text:span>d</text:span></text:p></table:table-cell>"#,
    ),
    (
        "span_tab_dropped",
        r#"<table:table-cell office:value-type="string"><text:p><text:span>a<text:tab/>b</text:span></text:p></table:table-cell>"#,
    ),
    (
        "span_line_break_dropped",
        r#"<table:table-cell office:value-type="string"><text:p><text:span>a<text:line-break/>b</text:span></text:p></table:table-cell>"#,
    ),
    (
        "link_in_span",
        r#"<table:table-cell office:value-type="string"><text:p><text:span><text:a xlink:href="x">q</text:a></text:span></text:p></table:table-cell>"#,
    ),
    (
        "span_in_link_only",
        r#"<table:table-cell office:value-type="string"><text:p><text:a xlink:href="x"><text:span>q</text:span></text:a></text:p></table:table-cell>"#,
    ),
    (
        "link_text_span_text",
        r#"<table:table-cell office:value-type="string"><text:p><text:a xlink:href="x">x<text:span>q</text:span>y</text:a></text:p></table:table-cell>"#,
    ),
    (
        "unknown_element_transparent",
        r#"<table:table-cell office:value-type="string"><text:p>a<text:foo>x<text:tab/>y</text:foo>b</text:p></table:table-cell>"#,
    ),
    (
        "span_in_unknown",
        r#"<table:table-cell office:value-type="string"><text:p><text:foo>a<text:span>q</text:span>b</text:foo></text:p></table:table-cell>"#,
    ),
    (
        "ruby_text_dropped",
        r#"<table:table-cell office:value-type="string"><text:p><text:ruby><text:ruby-base>base</text:ruby-base><text:ruby-text>top</text:ruby-text></text:ruby></text:p></table:table-cell>"#,
    ),
    (
        "tab_after_span",
        r#"<table:table-cell office:value-type="string"><text:p><text:span>a</text:span><text:tab/>b</text:p></table:table-cell>"#,
    ),
    (
        "paragraph_inside_nested_cell",
        r#"<table:table-cell office:value-type="string"><table:table-cell><text:p>no</text:p></table:table-cell><text:p>x</text:p></table:table-cell>"#,
    ),
];

fn string_entries() -> Vec<Crafted> {
    STRINGS
        .iter()
        .map(|(name, cell)| single(*name, cell))
        .collect()
}

/// Repeat counts of cells and rows, spelled in many ways.
fn repeat_entries() -> Vec<Crafted> {
    let mut entries = Vec::new();
    let row = |cells: &str| format!("<table:table-row>{cells}</table:table-row>");
    let at = |name: &str, rows: String| entry(name, rows, "A1:D4");
    for repeat in [
        "0",
        "-1",
        "x",
        "",
        "2",
        "3",
        "5x",
        " 5",
        "+5",
        "0x5",
        "5.0",
        "1e1",
        "2147483647",
        "2147483648",
    ] {
        entries.push(at(
            &format!("columns_repeated[{repeat}]"),
            row(&format!(
                r#"<table:table-cell office:value-type="float" office:value="2" table:number-columns-repeated="{repeat}"/>{}"#,
                float("3")
            )),
        ));
        entries.push(at(
            &format!("empty_columns_repeated[{repeat}]"),
            row(&format!(
                r#"<table:table-cell table:number-columns-repeated="{repeat}"/>{}"#,
                float("3")
            )),
        ));
        entries.push(at(
            &format!("rows_repeated[{repeat}]"),
            format!(
                r#"<table:table-row table:number-rows-repeated="{repeat}">{}</table:table-row>{}"#,
                float("1"),
                row(&float("2"))
            ),
        ));
        entries.push(at(
            &format!("empty_rows_repeated[{repeat}]"),
            format!(
                r#"<table:table-row table:number-rows-repeated="{repeat}"><table:table-cell/></table:table-row>{}"#,
                row(&float("2"))
            ),
        ));
    }
    entries
}

fn layout_entries() -> Vec<Crafted> {
    let mut entries = Vec::new();
    let row = |cells: &str| format!("<table:table-row>{cells}</table:table-row>");
    let at = |name: &str, rows: String| entry(name, rows, "A1:D4");
    entries.push(at(
        "covered_value_kept",
        row(r#"<table:table-cell office:value-type="float" office:value="1" table:number-columns-spanned="2"/><table:covered-table-cell office:value-type="float" office:value="9"/>"#),
    ));
    entries.push(at(
        "text_cell_repeated",
        row(r#"<table:table-cell office:value-type="string" table:number-columns-repeated="2"><text:p>t</text:p></table:table-cell>"#),
    ));
    entries.push(at(
        "covered_cell_repeated",
        row(&format!(
            r#"<table:covered-table-cell table:number-columns-repeated="2"/>{}"#,
            float("1")
        )),
    ));
    entries.push(at(
        "style_only_cell_advances",
        row(&format!(
            r#"<table:table-cell table:style-name="ce1"/>{}"#,
            float("1")
        )),
    ));
    for (name, wrapper) in [
        ("row_group", "table-row-group"),
        ("header_rows", "table-header-rows"),
        ("table_rows", "table-rows"),
    ] {
        entries.push(at(
            name,
            format!(
                "<table:{wrapper}>{}</table:{wrapper}>{}",
                row(&float("1")),
                row(&float("2"))
            ),
        ));
    }
    entries.push(at(
        "nested_row_groups",
        format!(
            "<table:table-row-group><table:table-row-group>{}</table:table-row-group></table:table-row-group>{}",
            row(&float("1")),
            row(&float("2"))
        ),
    ));
    entries.push(at(
        "header_rows_in_group",
        format!(
            "<table:table-row-group><table:table-header-rows>{}</table:table-header-rows></table:table-row-group>{}",
            row(&float("1")),
            row(&float("2"))
        ),
    ));
    entries
}

/// Elements that are ignored inside rows and cells, and those that change nothing.
fn ignored_structure_entries() -> Vec<Crafted> {
    let mut entries = Vec::new();
    let row = |cells: &str| format!("<table:table-row>{cells}</table:table-row>");
    let rows_of = |rows: &[String]| rows.concat();
    let at = |name: &str, rows: String| entry(name, rows, "A1:D4");
    entries.push(at(
        "row_in_row_ignored",
        rows_of(&[
            format!("<table:table-row>{}</table:table-row>", row(&float("1"))),
            row(&float("2")),
        ]),
    ));
    entries.push(at(
        "cell_in_cell_ignored",
        row(&format!(
            r#"<table:table-cell office:value-type="float" office:value="1"><table:table-cell office:value-type="float" office:value="5"/></table:table-cell>{}"#,
            float("2")
        )),
    ));
    entries.push(at(
        "columns_and_shapes_ignored",
        format!(
            r#"<table:table-column table:number-columns-repeated="3"/><table:shapes/>{}"#,
            row(&float("1"))
        ),
    ));
    entries.push(at(
        "row_with_style_only",
        format!(
            r#"<table:table-row table:style-name="ro1"><table:table-cell/></table:table-row>{}"#,
            row(&float("1"))
        ),
    ));
    entries.push(at(
        "collapsed_row",
        format!(
            r#"<table:table-row table:visibility="collapse">{}</table:table-row>"#,
            float("1")
        ),
    ));
    entries.push(at(
        "self_closing_row_repeated",
        format!(
            r#"<table:table-row table:number-rows-repeated="2"/>{}"#,
            row(&float("1"))
        ),
    ));
    entries.push(at(
        "later_table_leaves_the_first_alone",
        format!(
            r#"{}</table:table><table:table table:name="T2">{}"#,
            row(&float("1")),
            row(&float("2"))
        ),
    ));
    entries
}

fn formula_cell(formula: &str, rest: &str) -> String {
    format!(r#"<table:table-cell table:formula="{formula}" {rest}/>"#)
}

const FORMULAS: &[(&str, &str, &str)] = &[
    (
        "cached_number",
        "of:=1+1",
        r#"office:value-type="float" office:value="5""#,
    ),
    (
        "cached_zero",
        "of:=1+1",
        r#"office:value-type="float" office:value="0""#,
    ),
    ("cached_without_type", "of:=1+1", r#"office:value="5""#),
    ("no_cache", "of:=1+1", ""),
    (
        "plain_equals",
        "=1+1",
        r#"office:value-type="float" office:value="5""#,
    ),
    ("empty_formula", "of:=", ""),
    (
        "string_attribute_kept",
        "of:=&quot;x&quot;",
        r#"office:value-type="string" office:string-value="q""#,
    ),
    (
        "string_attribute_over_number_formula",
        "of:=1+1",
        r#"office:value-type="string" office:string-value="q""#,
    ),
    (
        "string_attribute_empty",
        "of:=&quot;x&quot;",
        r#"office:value-type="string" office:string-value="""#,
    ),
    (
        "boolean_cached",
        "of:=TRUE()",
        r#"office:value-type="boolean" office:boolean-value="false""#,
    ),
    (
        "date_cached",
        "of:=DATE(2020;1;2)",
        r#"office:value-type="date" office:date-value="2000-01-01""#,
    ),
    (
        "error_cached",
        "of:=1/0",
        r#"office:value-type="string" office:string-value="" calcext:value-type="error""#,
    ),
    (
        "volatile",
        "of:=NOW()",
        r#"office:value-type="float" office:value="5""#,
    ),
    (
        "old_namespace",
        "oooc:=1+1",
        r#"office:value-type="float" office:value="5""#,
    ),
    (
        "unknown_namespace",
        "foo:=1+1",
        r#"office:value-type="float" office:value="5""#,
    ),
    (
        "invalid_text",
        "of:=1+",
        r#"office:value-type="float" office:value="5""#,
    ),
    (
        "self_reference",
        "of:=[.A1]",
        r#"office:value-type="float" office:value="5""#,
    ),
    (
        "repeat_ignored",
        "of:=1+1",
        r#"table:number-columns-repeated="3""#,
    ),
];

fn formula_entries() -> Vec<Crafted> {
    let mut entries: Vec<Crafted> = FORMULAS
        .iter()
        .map(|(name, formula, rest)| single(*name, &formula_cell(formula, rest)))
        .collect();
    entries.push(entry(
        "reference",
        format!(
            "<table:table-row>{}{}</table:table-row>",
            formula_cell(
                "of:=[.B1]*2",
                r#"office:value-type="float" office:value="100""#
            ),
            float("3")
        ),
        "A1:D1",
    ));
    entries.push(entry(
        "repeat_ignored_then_cell",
        format!(
            "<table:table-row>{}{}</table:table-row>",
            formula_cell("of:=1+1", r#"table:number-columns-repeated="2""#),
            float("7")
        ),
        "A1:D1",
    ));
    entries.push(entry(
        "covered_formula",
        "<table:table-row><table:table-cell/><table:covered-table-cell table:formula=\"of:=1+1\"/></table:table-row>",
        "A1:D1",
    ));
    entries
}

fn table_style(name: &str, family: &str, display: &str) -> String {
    format!(
        r#"<style:style style:name="{name}" style:family="{family}"><style:table-properties table:display="{display}"/></style:style>"#
    )
}

/// Documents of three tables, `S1` to `S3`, whose styles hide some of them.
fn visibility_entries() -> Vec<Crafted> {
    let cell = |number: u32| {
        format!(
            r#"<table:table-row><table:table-cell office:value-type="float" office:value="{number}"/></table:table-row>"#
        )
    };
    let tables = |attributes: [&str; 3]| -> String {
        let mut tables = String::new();
        for (number, attribute) in (1_u32..).zip(attributes) {
            let table = format!(
                r#"<table:table table:name="S{number}" {attribute}>{}</table:table>"#,
                cell(number)
            );
            tables.push_str(&table);
        }
        tables
    };
    let hidden = table_style("hid", "table", "false");
    let shown = table_style("vis", "table", "true");
    let mut documents = vec![
        (
            "second_hidden",
            hidden.clone(),
            tables(["", r#"table:style-name="hid""#, ""]),
        ),
        (
            "first_hidden",
            hidden.clone(),
            tables([r#"table:style-name="hid""#, "", ""]),
        ),
        (
            "last_hidden",
            hidden.clone(),
            tables(["", "", r#"table:style-name="hid""#]),
        ),
        (
            "all_but_first_hidden",
            hidden.clone(),
            tables(["", r#"table:style-name="hid""#, r#"table:style-name="hid""#]),
        ),
        (
            "shown_style",
            shown.clone(),
            tables(["", r#"table:style-name="vis""#, ""]),
        ),
        (
            "missing_style",
            hidden.clone(),
            tables(["", r#"table:style-name="nothing""#, ""]),
        ),
        (
            "cell_style_family",
            table_style("hid", "table-cell", "false"),
            tables(["", r#"table:style-name="hid""#, ""]),
        ),
        (
            "display_on_the_table",
            String::new(),
            tables(["", r#"table:display="false""#, ""]),
        ),
        (
            "display_false_in_other_case",
            table_style("hid", "table", "False"),
            tables(["", r#"table:style-name="hid""#, ""]),
        ),
        (
            "display_zero",
            table_style("hid", "table", "0"),
            tables(["", r#"table:style-name="hid""#, ""]),
        ),
        (
            "display_with_space",
            table_style("hid", "table", "false "),
            tables(["", r#"table:style-name="hid""#, ""]),
        ),
        (
            "display_empty",
            table_style("hid", "table", ""),
            tables(["", r#"table:style-name="hid""#, ""]),
        ),
        (
            "two_styles_same_name",
            format!("{shown}{hidden}").replace(r#"style:name="vis""#, r#"style:name="hid""#),
            tables(["", r#"table:style-name="hid""#, ""]),
        ),
    ];
    documents
        .drain(..)
        .map(|(name, styles, body)| Crafted {
            raw: Some(document_xml(&styles, &body)),
            ..entry(name, "", "A1:A1")
        })
        .collect()
}

const NULL_DATES: &[&str] = &[
    "1899-12-30",
    "1904-01-01",
    "1900-01-01",
    "2000-01-01",
    "1899-12-31",
    "abc",
    "",
    " 1904-01-01",
    "1904-01-01T12:00:00",
    "1970-01-01",
    "0001-01-01",
    "-0001-01-01",
    "2020-01-02",
    "1904-1-1",
    "1904-01-01Z",
    "9999-12-31",
    "1900-02-29",
    "0000-01-01",
];

/// Date cells under the null date a document names.
fn date_system_entries() -> Vec<Crafted> {
    let mut entries = Vec::new();
    for null in NULL_DATES {
        for date in [
            "2020-01-02",
            "1904-01-01",
            "1900-03-01",
            "1899-12-30",
            "2020-01-02T12:00:00",
        ] {
            entries.push(Crafted {
                name: format!("null_date[{null}] date[{date}]"),
                rows: format!(
                    r#"<table:table-row><table:table-cell office:value-type="date" office:date-value="{date}"/></table:table-row>"#
                ),
                settings: format!(
                    r#"<table:calculation-settings><table:null-date table:value-type="date" table:date-value="{null}"/></table:calculation-settings>"#
                ),
                range: "A1:A1",
                renamed: false,
                raw: None,
            });
        }
    }
    entries
}

/// Repeats and matrices at the edge of the sheet, read where the edge is.
fn edge_entries() -> Vec<Crafted> {
    let row = |cells: &str| format!("<table:table-row>{cells}</table:table-row>");
    let empty =
        |count: u32| format!(r#"<table:table-cell table:number-columns-repeated="{count}"/>"#);
    let matrix = |columns: u32, rows: u32| {
        format!(
            r#"<table:table-cell table:formula="of:=1+1" table:number-matrix-columns-spanned="{columns}" table:number-matrix-rows-spanned="{rows}"/>"#
        )
    };
    let tall = |rows: u32, cell: &str| {
        format!(
            r#"<table:table-row table:number-rows-repeated="{rows}"><table:table-cell/></table:table-row><table:table-row>{cell}</table:table-row>"#
        )
    };
    let mut entries = vec![
        entry(
            "columns_repeated_past_the_edge",
            row(&format!(
                r#"{}<table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="5"/>"#,
                empty(16382)
            )),
            "XFB1:XFD1",
        ),
        entry(
            "rows_repeated_past_the_edge",
            format!(
                r#"<table:table-row table:number-rows-repeated="1048574"><table:table-cell/></table:table-row><table:table-row table:number-rows-repeated="5">{}</table:table-row>"#,
                float("1")
            ),
            "A1048574:A1048576",
        ),
    ];
    for (name, columns, rows) in [
        ("last_column", 1, 1),
        ("past_last_column", 2, 1),
        ("two_rows", 1, 2),
    ] {
        entries.push(entry(
            format!("matrix_at_the_edge[{name}]"),
            row(&format!("{}{}", empty(16383), matrix(columns, rows))),
            "XFC1:XFD2",
        ));
    }
    for (name, rows) in [("last_row", 1), ("past_last_row", 2)] {
        entries.push(entry(
            format!("matrix_at_the_edge[{name}]"),
            tall(1_048_575, &matrix(1, rows)),
            "A1048575:B1048576",
        ));
    }
    entries
}

/// Documents that are not well-formed, which `LibreOffice` refuses to load.
fn malformed_entries() -> Vec<Crafted> {
    let rows = r#"<table:table-row><table:table-cell office:value-type="float" office:value="1"/><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell office:value-type="float" office:value="2"/></table:table-row>"#;
    let full = content_xml(rows, "");
    let second_row =
        r#"<table:table-row><table:table-cell office:value-type="float" office:value="2""#;
    let cut_at =
        |marker: &str, keep: usize| full[..full.find(marker).unwrap_or(0) + keep].to_owned();
    let documents = [
        ("well_formed", full.clone()),
        ("table_not_closed", full.replace("</table:table>", "")),
        (
            "root_not_closed",
            full.replace("</office:document-content>", ""),
        ),
        ("cut_after_a_row", cut_at(second_row, 0)),
        ("cut_inside_a_tag", cut_at("office:value=\"2\"", 0)),
        ("cut_inside_a_paragraph", cut_at("a</text:p>", 1)),
        (
            "mismatched_end_tag",
            full.replace("a</text:p>", "a</text:q>"),
        ),
        (
            "extra_end_tag",
            full.replace("a</text:p>", "a</text:p></text:p>"),
        ),
        (
            "stray_end_tag_before_the_table",
            full.replace(
                "<table:table table:name",
                "</table:foo><table:table table:name",
            ),
        ),
        (
            "undefined_entity",
            full.replace("<text:p>a", "<text:p>&bogus;a"),
        ),
        ("two_roots", format!("{full}{full}")),
        ("unfinished_tag_after_the_root", format!("{full}<junk")),
        ("text_after_the_root", format!("{full} text")),
        ("whitespace_after_the_root", format!("{full}\n  \n")),
    ];
    documents
        .into_iter()
        .map(|(name, document)| Crafted {
            raw: Some(document),
            ..entry(name, "", "A1:B2")
        })
        .collect()
}

const MATRICES: &[(&str, &str, &str)] = &[
    ("wide", "3", "1"),
    ("single", "1", "1"),
    ("square", "2", "2"),
    ("zero_columns", "0", "1"),
    ("negative", "-1", "1"),
    ("text", "x", "1"),
    ("huge_rows", "3", "99999999999"),
    ("past_columns", "99999", "1"),
    ("prefix", "2x", " 1"),
    ("tall", "1", "4"),
    ("empty_columns", "", "1"),
    ("all_columns", "16384", "1"),
    ("past_all_columns", "16385", "1"),
    ("all_rows", "1", "1048576"),
    ("past_all_rows", "1", "1048577"),
    ("columns_20000", "20000", "1"),
    ("columns_32767", "32767", "1"),
    ("columns_32768", "32768", "1"),
    ("columns_65535", "65535", "1"),
    ("columns_65536", "65536", "1"),
    ("columns_99999", "99999", "1"),
    ("rows_1100000", "1", "1100000"),
    ("rows_2000000", "1", "2000000"),
    ("rows_16777215", "1", "16777215"),
    ("rows_16777216", "1", "16777216"),
    ("rows_99999999", "1", "99999999"),
    ("rows_2147483647", "1", "2147483647"),
];

fn matrix_entries() -> Vec<Crafted> {
    let mut entries = Vec::new();
    for (name, columns, rows) in MATRICES {
        entries.push(entry(
            format!("matrix_{name}"),
            format!(
                r#"<table:table-row><table:table-cell table:formula="of:=1+1" table:number-matrix-columns-spanned="{columns}" table:number-matrix-rows-spanned="{rows}"/></table:table-row>"#
            ),
            "A1:D5",
        ));
    }
    for (name, cell) in [
        (
            "matrix_only_columns",
            r#"<table:table-cell table:formula="of:=1+1" table:number-matrix-columns-spanned="2"/>"#,
        ),
        (
            "matrix_only_rows",
            r#"<table:table-cell table:formula="of:=1+1" table:number-matrix-rows-spanned="2"/>"#,
        ),
        (
            "matrix_without_formula",
            r#"<table:table-cell office:value-type="float" office:value="4" table:number-matrix-columns-spanned="2" table:number-matrix-rows-spanned="1"/>"#,
        ),
        (
            "spanned_is_not_matrix",
            r#"<table:table-cell table:formula="of:=1+1" table:number-columns-spanned="2"/><table:covered-table-cell/>"#,
        ),
    ] {
        entries.push(single(name, cell));
    }
    entries.push(entry(
        "matrix_over_cell",
        format!(
            r#"<table:table-row><table:table-cell table:formula="of:=1+1" table:number-matrix-columns-spanned="2" table:number-matrix-rows-spanned="1"/>{}</table:table-row>"#,
            float("9")
        ),
        "A1:D2",
    ));
    entries.push(entry(
        "matrix_covered_values",
        r#"<table:table-row><table:table-cell table:formula="of:=1+1" table:number-matrix-columns-spanned="2" table:number-matrix-rows-spanned="2"/><table:covered-table-cell office:value-type="float" office:value="4"/></table:table-row><table:table-row><table:covered-table-cell office:value-type="float" office:value="5"/><table:covered-table-cell office:value-type="float" office:value="6"/></table:table-row>"#,
        "A1:D3",
    ));
    entries
}

/// Cell values of hand-written tables read as the oracle gives them.
#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn typed_cell_values_over_ods() -> Result<(), Box<dyn StdError>> {
    let entries = [scalar_entries(), date_entries(), typing_entries()].concat();
    assert_reads_as_the_oracle("scalar", &entries)
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn repeats_and_matrices_at_the_edge_of_the_sheet_over_ods() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("edge", &edge_entries())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn documents_that_are_not_well_formed_over_ods() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("malformed", &malformed_entries())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn hidden_tables_over_ods() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("visibility", &visibility_entries())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn namespaces_are_resolved_by_uri_over_ods() -> Result<(), Box<dyn StdError>> {
    let entries: Vec<Crafted> = [string_entries(), formula_entries(), repeat_entries()]
        .concat()
        .into_iter()
        .step_by(5)
        .map(|entry| Crafted {
            renamed: true,
            ..entry
        })
        .collect();
    assert_reads_as_the_oracle("renamed", &entries)
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn dates_follow_the_null_date_over_ods() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("date_system", &date_system_entries())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn text_of_string_cells_over_ods() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("string", &string_entries())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn rows_columns_and_repeats_over_ods() -> Result<(), Box<dyn StdError>> {
    let entries = [
        layout_entries(),
        ignored_structure_entries(),
        repeat_entries(),
    ]
    .concat();
    assert_reads_as_the_oracle("layout", &entries)
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn formula_cells_over_ods() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("formula", &formula_entries())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn matrix_formulas_fill_their_area_over_ods() -> Result<(), Box<dyn StdError>> {
    assert_reads_as_the_oracle("matrix", &matrix_entries())
}
