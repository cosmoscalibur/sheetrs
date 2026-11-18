use std::error::Error as StdError;
use std::io::Cursor;

use sheetrs_core::{
    Cell, CellPos, DateSystem, ErrorKind, Formula, SheetView, Value, Visibility, Workbook,
};
use sheetrs_xlsx::{Error, read};

mod common;
use common::{package, rels_xml, single_sheet_package_with, workbook_xml, worksheet_xml};

type TestResult = Result<(), Box<dyn StdError>>;

fn single_sheet_package(
    rows: &str,
    shared_strings: Option<&str>,
) -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    single_sheet_package_with(rows, shared_strings, "")
}

/// A one-sheet package read back into a workbook.
fn single_sheet(rows: &str, shared_strings: Option<&str>) -> Result<Workbook, Box<dyn StdError>> {
    Ok(read(single_sheet_package(rows, shared_strings)?)?)
}

fn sheet_of(workbook: &Workbook) -> Result<SheetView<'_>, Box<dyn StdError>> {
    Ok(workbook.sheet("S1").ok_or("sheet S1 is missing")?)
}

/// The cell at `pos`, owned.
fn cell_at(sheet: &SheetView<'_>, pos: CellPos) -> Option<Cell> {
    sheet.get(pos).map(|view| view.to_cell())
}

fn at(row: u32, col: u32) -> CellPos {
    CellPos { row, col }
}

fn constant(value: Value) -> Cell {
    Cell {
        value,
        formula: None,
        stale: false,
    }
}

/// A cell that stands for a formula whose result is stored, as a boolean or an error name.
fn implied(value: Value, formula: &str) -> Cell {
    Cell {
        value,
        formula: Some(Formula::Text(formula.to_owned())),
        stale: true,
    }
}

fn text(content: &str) -> Value {
    Value::Text(content.into())
}

#[test]
fn sheets_keep_order_names_visibility_and_date_system() -> TestResult {
    let empty_sheet = worksheet_xml("");
    let source = package(&[
        (
            "xl/workbook.xml",
            workbook_xml(
                r#"<sheet name="S1" sheetId="1" r:id="rId1"/><sheet name="S2" sheetId="2" state="hidden" r:id="rId2"/><sheet name="S3" sheetId="3" state="veryHidden" r:id="rId3"/>"#,
                r#"<workbookPr date1904="1"/>"#,
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            rels_xml(
                r#"<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="worksheet" Target="/xl/worksheets/sheet2.xml"/><Relationship Id="rId3" Type="worksheet" Target="worksheets/sheet3.xml"/>"#,
            ),
        ),
        ("xl/worksheets/sheet1.xml", empty_sheet.clone()),
        ("xl/worksheets/sheet2.xml", empty_sheet.clone()),
        ("xl/worksheets/sheet3.xml", empty_sheet),
    ])?;
    let workbook = read(source)?;
    let summary: Vec<(&str, Visibility)> = workbook
        .sheets()
        .map(|sheet| (sheet.name(), sheet.visibility()))
        .collect();
    assert_eq!(
        summary,
        vec![
            ("S1", Visibility::Visible),
            ("S2", Visibility::Hidden),
            ("S3", Visibility::VeryHidden),
        ]
    );
    assert_eq!(workbook.date_system(), DateSystem::Date1904);
    Ok(())
}

#[test]
fn the_1900_date_system_is_the_default() -> TestResult {
    let workbook = single_sheet("", None)?;
    assert_eq!(workbook.date_system(), DateSystem::Date1900);
    Ok(())
}

#[test]
fn cell_types_are_read() -> TestResult {
    let workbook = single_sheet(
        r#"<row r="1"><c r="A1"><v>1.5</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="b"><v>1</v></c><c r="D1" t="e"><v>#DIV/0!</v></c><c r="E1" t="inlineStr"><is><t>inline</t></is></c></row><row r="2"><c r="A2" t="str"><f>"x"&amp;"y"</f><v>xy</v></c><c r="B2"><f>A1*2</f><v>3</v></c><c r="C2" s="3"/><c r="D2"><v>0.30000000000000004</v></c></row>"#,
        Some("<si><t>hello</t></si>"),
    )?;
    let sheet = sheet_of(&workbook)?;
    assert_eq!(
        cell_at(&sheet, at(0, 0)),
        Some(constant(Value::Number(1.5)))
    );
    assert_eq!(cell_at(&sheet, at(0, 1)), Some(constant(text("hello"))));
    assert_eq!(
        cell_at(&sheet, at(0, 2)),
        Some(implied(Value::Bool(true), "TRUE()"))
    );
    assert_eq!(
        cell_at(&sheet, at(0, 3)),
        Some(implied(Value::Error(ErrorKind::DivZero), "#DIV/0!"))
    );
    assert_eq!(cell_at(&sheet, at(0, 4)), Some(constant(text("inline"))));
    assert_eq!(
        cell_at(&sheet, at(1, 0)),
        Some(Cell {
            value: text("xy"),
            formula: Some(Formula::Text(r#""x"&"y""#.to_owned())),
            stale: true,
        })
    );
    assert_eq!(
        cell_at(&sheet, at(1, 1)),
        Some(Cell {
            value: Value::Number(3.0),
            formula: Some(Formula::Text("A1*2".to_owned())),
            stale: false,
        })
    );
    assert_eq!(
        cell_at(&sheet, at(1, 2)),
        None,
        "a style-only cell holds nothing"
    );
    let Some(Cell {
        value: Value::Number(sum),
        ..
    }) = cell_at(&sheet, at(1, 3))
    else {
        return Err("D2 is not a number".into());
    };
    assert_eq!(sum.to_bits(), (0.1_f64 + 0.2).to_bits());
    Ok(())
}

#[test]
fn shared_strings_join_rich_text_and_skip_phonetic_runs() -> TestResult {
    let workbook = single_sheet(
        r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row>"#,
        Some(
            r#"<si><t>hello</t></si><si><r><t>ab</t></r><r><rPr/><t xml:space="preserve"> cd</t></r></si><si><t>x</t><rPh sb="0" eb="1"><t>ignored</t></rPh></si><si><t>a &amp; b &#65;</t></si>"#,
        ),
    )?;
    let sheet = sheet_of(&workbook)?;
    let strings: Vec<Option<Value>> = (0..4)
        .map(|col| cell_at(&sheet, at(0, col)).map(|cell| cell.value))
        .collect();
    assert_eq!(
        strings,
        vec![
            Some(text("hello")),
            Some(text("ab cd")),
            Some(text("x")),
            Some(text("a & b A")),
        ]
    );
    Ok(())
}

#[test]
fn shared_formulas_record_their_anchor() -> TestResult {
    let workbook = single_sheet(
        r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>3</v></c></row>"#,
        None,
    )?;
    let sheet = sheet_of(&workbook)?;
    assert_eq!(
        cell_at(&sheet, at(0, 0)).and_then(|cell| cell.formula.clone()),
        Some(Formula::Text("B1+1".to_owned()))
    );
    assert_eq!(
        cell_at(&sheet, at(1, 0)).and_then(|cell| cell.formula.clone()),
        Some(Formula::Shared { index: 0 })
    );
    assert_eq!(sheet.shared_anchor(0), Some(at(0, 0)));
    Ok(())
}

#[test]
fn cells_without_addresses_follow_the_previous_ones() -> TestResult {
    let workbook = single_sheet(
        "<row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row>",
        None,
    )?;
    let sheet = sheet_of(&workbook)?;
    let numbers: Vec<(u32, u32, Value)> = sheet
        .cells()
        .map(|(pos, cell)| (pos.row, pos.col, cell.to_cell().value))
        .collect();
    assert_eq!(
        numbers,
        vec![
            (0, 0, Value::Number(1.0)),
            (0, 1, Value::Number(2.0)),
            (1, 0, Value::Number(3.0)),
        ]
    );
    Ok(())
}

#[test]
fn a_missing_worksheet_part_is_reported_by_name() -> TestResult {
    let source = package(&[
        (
            "xl/workbook.xml",
            workbook_xml(r#"<sheet name="S1" sheetId="1" r:id="rId1"/>"#, ""),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            rels_xml(r#"<Relationship Id="rId1" Type="worksheet" Target="worksheets/gone.xml"/>"#),
        ),
    ])?;
    assert!(matches!(
        read(source),
        Err(Error::MissingPart(part)) if part == "xl/worksheets/gone.xml"
    ));
    Ok(())
}

#[test]
fn input_that_is_not_a_zip_package_is_rejected() {
    assert!(matches!(
        read(Cursor::new(b"not a zip".to_vec())),
        Err(Error::Zip(_))
    ));
}

#[test]
fn a_shared_string_index_out_of_range_holds_nothing() -> TestResult {
    let cell = only_cell(
        &cell_xml(r#" t="s""#, "<v>5</v>"),
        Some("<si><t>a</t></si>"),
    )?;
    assert_eq!(cell, None);
    Ok(())
}
fn only_cell(rows: &str, shared_strings: Option<&str>) -> Result<Option<Cell>, Box<dyn StdError>> {
    let workbook = single_sheet(rows, shared_strings)?;
    Ok(cell_at(&sheet_of(&workbook)?, at(0, 0)))
}

fn cell_xml(attributes: &str, children: &str) -> String {
    format!(r#"<row r="1"><c r="A1"{attributes}>{children}</c></row>"#)
}

#[test]
fn a_formula_with_an_empty_cached_value_keeps_the_formula() -> TestResult {
    let cell = only_cell(&cell_xml("", "<f>1+1</f><v></v>"), None)?;
    assert_eq!(
        cell,
        Some(Cell {
            value: Value::Empty,
            formula: Some(Formula::Text("1+1".to_owned())),
            stale: true,
        })
    );
    Ok(())
}

#[test]
fn an_inline_string_cell_without_content_holds_nothing() -> TestResult {
    let cell = only_cell(&cell_xml(r#" t="inlineStr""#, "<v>abc</v>"), None)?;
    assert_eq!(cell, None);
    Ok(())
}

#[test]
fn booleans_follow_the_numeric_and_word_forms() -> TestResult {
    for (stored, expected) in [
        ("0", false),
        ("1", true),
        ("2", true),
        ("true", true),
        ("TRUE", true),
        ("false", false),
        ("FALSE", false),
    ] {
        let cell = only_cell(&cell_xml(r#" t="b""#, &format!("<v>{stored}</v>")), None)?;
        let formula = if expected { "TRUE()" } else { "FALSE()" };
        assert_eq!(
            cell,
            Some(implied(Value::Bool(expected), formula)),
            "{stored}"
        );
    }
    Ok(())
}

#[test]
fn text_in_a_boolean_cell_is_read_as_libreoffice_reads_it() -> TestResult {
    for stored in ["maybe", "yes", "T", " "] {
        let cell = only_cell(&cell_xml(r#" t="b""#, &format!("<v>{stored}</v>")), None)?;
        assert_eq!(
            cell,
            Some(implied(Value::Bool(false), "FALSE()")),
            "{stored:?}"
        );
    }
    Ok(())
}
#[test]
fn escapes_are_decoded_in_shared_and_inline_strings_only() -> TestResult {
    let shared = only_cell(
        &cell_xml(r#" t="s""#, "<v>0</v>"),
        Some("<si><t>a_x000D_b_x005F_x</t></si>"),
    )?;
    assert_eq!(shared, Some(constant(text("a\rb_x"))));
    let inline = only_cell(
        &cell_xml(r#" t="inlineStr""#, "<is><t>_x0041__x0042_</t></is>"),
        None,
    )?;
    assert_eq!(inline, Some(constant(text("AB"))));
    let result = only_cell(&cell_xml(r#" t="str""#, "<v>x_x000A_y</v>"), None)?;
    assert_eq!(result, Some(constant(text("x_x000A_y"))));
    Ok(())
}

#[test]
fn text_that_only_looks_like_an_escape_is_kept() -> TestResult {
    let kept = only_cell(
        &cell_xml(r#" t="inlineStr""#, "<is><t>_x00ZZ_ _x41_ _x0041</t></is>"),
        None,
    )?;
    assert_eq!(kept, Some(constant(text("_x00ZZ_ _x41_ _x0041"))));
    Ok(())
}

#[test]
fn numbers_ignore_surrounding_whitespace() -> TestResult {
    let cell = only_cell(&cell_xml("", "<v> 1 </v>"), None)?;
    assert_eq!(cell, Some(constant(Value::Number(1.0))));
    Ok(())
}

#[test]
fn text_in_a_numeric_cell_is_read_as_libreoffice_reads_it() -> TestResult {
    for (stored, expected) in [("nan", 0.0), ("abc", 0.0), ("1e999", f64::INFINITY)] {
        let cell = only_cell(&cell_xml("", &format!("<v>{stored}</v>")), None)?;
        assert_eq!(cell, Some(constant(Value::Number(expected))), "{stored:?}");
    }
    Ok(())
}
#[test]
fn date_typed_cells_become_serial_numbers_or_stay_text() -> TestResult {
    let date = only_cell(&cell_xml(r#" t="d""#, "<v>2020-01-02</v>"), None)?;
    assert_eq!(date, Some(constant(Value::Number(43_832.0))));
    let not_a_date = only_cell(&cell_xml(r#" t="d""#, "<v>abc</v>"), None)?;
    assert_eq!(not_a_date, Some(constant(text("abc"))));
    Ok(())
}
fn truncated_worksheet(content: &str) -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    package(&[
        (
            "xl/workbook.xml",
            workbook_xml(r#"<sheet name="S1" sheetId="1" r:id="rId1"/>"#, ""),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            rels_xml(
                r#"<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>"#,
            ),
        ),
        ("xl/worksheets/sheet1.xml", content.to_owned()),
    ])
}

#[test]
fn a_worksheet_cut_at_a_tag_boundary_keeps_the_cells_read_over_xlsx() -> TestResult {
    let source = truncated_worksheet(r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c>"#)?;
    let workbook = read(source)?;
    assert_eq!(
        cell_at(&sheet_of(&workbook)?, at(0, 0)),
        Some(constant(Value::Number(1.0)))
    );
    Ok(())
}

#[test]
fn a_cell_whose_address_is_not_a_cell_goes_where_an_unaddressed_one_would() -> TestResult {
    let workbook = single_sheet(
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="XFE1"><v>2</v></c><c><v>3</v></c></row>"#,
        None,
    )?;
    let numbers: Vec<(u32, u32, Value)> = sheet_of(&workbook)?
        .cells()
        .map(|(pos, cell)| (pos.row, pos.col, cell.to_cell().value))
        .collect();
    assert_eq!(
        numbers,
        vec![
            (0, 0, Value::Number(1.0)),
            (0, 1, Value::Number(2.0)),
            (0, 2, Value::Number(3.0)),
        ]
    );
    Ok(())
}

#[test]
fn a_row_outside_the_sheet_drops_its_unaddressed_cells_and_the_next_row_follows_it() -> TestResult {
    let past_the_end = single_sheet(
        r#"<row r="1"><c><v>1</v></c></row><row r="1048577"><c><v>2</v></c></row><row><c><v>3</v></c></row>"#,
        None,
    )?;
    assert_eq!(sheet_of(&past_the_end)?.len(), 1);
    let row_zero = single_sheet(
        r#"<row r="2"><c><v>1</v></c></row><row r="0"><c><v>2</v></c></row><row><c><v>3</v></c></row>"#,
        None,
    )?;
    let sheet = sheet_of(&row_zero)?;
    assert_eq!(
        cell_at(&sheet, at(1, 0)),
        Some(constant(Value::Number(1.0)))
    );
    assert_eq!(
        cell_at(&sheet, at(0, 0)),
        Some(constant(Value::Number(3.0)))
    );
    assert_eq!(sheet.len(), 2);
    Ok(())
}

#[test]
fn an_addressed_cell_keeps_its_place_in_a_row_outside_the_sheet() -> TestResult {
    let workbook = single_sheet(r#"<row r="1048577"><c r="B2"><v>2</v></c></row>"#, None)?;
    assert_eq!(
        cell_at(&sheet_of(&workbook)?, at(1, 1)),
        Some(constant(Value::Number(2.0)))
    );
    Ok(())
}

#[test]
fn an_empty_value_holds_nothing_whatever_the_cell_type() -> TestResult {
    for kind in [
        "",
        r#" t="n""#,
        r#" t="s""#,
        r#" t="b""#,
        r#" t="str""#,
        r#" t="e""#,
    ] {
        for value in ["<v></v>", "<v/>"] {
            let cell = only_cell(&cell_xml(kind, value), Some("<si><t>a</t></si>"))?;
            assert_eq!(cell, None, "{kind} {value}");
        }
    }
    Ok(())
}

#[test]
fn inline_content_counts_only_for_inline_string_cells() -> TestResult {
    let shared = Some("<si><t>a</t></si>");
    let numeric = only_cell(&cell_xml(r#" t="n""#, "<is><t>x</t></is>"), shared)?;
    assert_eq!(numeric, None);
    let untyped = only_cell(&cell_xml("", "<is><t>x</t></is>"), shared)?;
    assert_eq!(untyped, None);
    let shared_cell = only_cell(&cell_xml(r#" t="s""#, "<v>0</v><is><t>x</t></is>"), shared)?;
    assert_eq!(shared_cell, Some(constant(text("a"))));
    let empty_inline = only_cell(&cell_xml(r#" t="inlineStr""#, "<is></is>"), shared)?;
    assert_eq!(empty_inline, None);
    Ok(())
}

#[test]
fn an_empty_shared_string_reads_as_an_empty_cell() -> TestResult {
    for strings in ["<si/>", "<si><t></t></si>"] {
        let cell = only_cell(&cell_xml(r#" t="s""#, "<v>0</v>"), Some(strings))?;
        assert_eq!(cell, None, "{strings}");
    }
    Ok(())
}

#[test]
fn escaped_surrogate_pairs_combine_and_a_lone_surrogate_ends_the_text() -> TestResult {
    let cell = only_cell(
        &cell_xml(
            r#" t="inlineStr""#,
            "<is><t>_xD83D__xDE00_ _xD83D_ _xDE00_</t></is>",
        ),
        None,
    )?;
    assert_eq!(cell, Some(constant(text("\u{1F600} "))));
    Ok(())
}

#[test]
fn leading_whitespace_is_skipped_but_not_a_non_breaking_space() -> TestResult {
    let spaced = only_cell(
        &cell_xml(
            "",
            "<v>	 1
</v>",
        ),
        None,
    )?;
    assert_eq!(spaced, Some(constant(Value::Number(1.0))));
    let non_breaking = only_cell(&cell_xml("", "<v>&#160;1</v>"), None)?;
    assert_eq!(non_breaking, Some(constant(Value::Number(0.0))));
    Ok(())
}
#[test]
fn an_empty_workbook_root_reads_as_a_workbook_without_sheets() -> TestResult {
    let source = package(&[
        ("xl/workbook.xml", "<workbook/>".to_owned()),
        ("xl/_rels/workbook.xml.rels", rels_xml("")),
    ])?;
    let workbook = read(source)?;
    assert!(workbook.sheets().next().is_none());
    Ok(())
}

#[test]
fn a_row_numbered_minus_one_has_no_row_number() -> TestResult {
    let workbook = single_sheet(
        r#"<row r="3x"><c><v>1</v></c></row><row r="-1"><c><v>2</v></c></row><row><c><v>3</v></c></row>"#,
        None,
    )?;
    let sheet = sheet_of(&workbook)?;
    assert_eq!(
        cell_at(&sheet, at(2, 0)),
        Some(constant(Value::Number(1.0)))
    );
    assert_eq!(
        cell_at(&sheet, at(3, 0)),
        Some(constant(Value::Number(2.0)))
    );
    assert_eq!(
        cell_at(&sheet, at(4, 0)),
        Some(constant(Value::Number(3.0)))
    );
    Ok(())
}

fn formula_text_at(sheet: &SheetView<'_>, row: u32, col: u32) -> Option<Formula> {
    cell_at(sheet, at(row, col)).and_then(|cell| cell.formula)
}

#[test]
fn an_array_formula_fills_its_area_and_replaces_cached_cells() -> TestResult {
    let workbook = single_sheet(
        r#"<row r="1"><c r="A1"><f t="array" ref="A1:B2">1+1</f><v>7</v></c><c r="B1"><v>5</v></c></row><row r="2"><c r="A2"><f>3+3</f><v>6</v></c></row>"#,
        None,
    )?;
    let sheet = sheet_of(&workbook)?;
    assert_eq!(sheet.len(), 4);
    let array = Some(Formula::Text("1+1".to_owned()));
    for (row, col) in [(0, 0), (0, 1), (1, 1)] {
        assert_eq!(formula_text_at(&sheet, row, col), array);
    }
    // A cell that holds a formula is replaced by another formula, never by a constant.
    assert_eq!(
        formula_text_at(&sheet, 1, 0),
        Some(Formula::Text("3+3".to_owned()))
    );
    Ok(())
}

#[test]
fn an_array_formula_whose_ref_does_not_start_at_its_cell_is_dropped() -> TestResult {
    for reference in [r#" ref="A2:A3""#, r#" ref="zz""#, ""] {
        let workbook = single_sheet(
            &format!(r#"<row r="1"><c r="A1"><f t="array"{reference}>1+1</f><v>7</v></c></row>"#),
            None,
        )?;
        assert_eq!(sheet_of(&workbook)?.len(), 0, "{reference}");
    }
    Ok(())
}

#[test]
fn an_array_formula_with_an_open_end_fills_its_first_row() -> TestResult {
    let workbook = single_sheet(
        r#"<row r="1"><c r="A1"><f t="array" ref="A1:B">1+1</f><v>7</v></c></row>"#,
        None,
    )?;
    let sheet = sheet_of(&workbook)?;
    assert_eq!(sheet.len(), 2);
    assert!(formula_text_at(&sheet, 0, 1).is_some());
    Ok(())
}

#[test]
fn an_array_formula_that_covers_too_many_cells_is_refused() -> TestResult {
    let rows = r#"<row r="1"><c r="A1"><f t="array" ref="A1:XFD1048576">1+1</f></c></row>"#;
    let package = single_sheet_package_with(rows, None, "")?;
    assert!(matches!(
        read(package),
        Err(Error::TooLarge(sheetrs_core::LimitError::ArrayCells))
    ));
    Ok(())
}

#[test]
fn boolean_and_error_cells_without_formula_carry_the_formula_they_stand_for() -> TestResult {
    for (kind, stored, formula) in [
        ("b", "1", Some("TRUE()")),
        ("b", "0", Some("FALSE()")),
        ("e", "#N/A", Some("#N/A")),
        ("e", " #div/0! ", Some("#DIV/0!")),
        ("e", "#SPILL!", Some("#SPILL!")),
        ("e", "x", None),
        ("e", "#GETTING_DATA", None),
        ("e", "#N/A!", None),
    ] {
        let cell = only_cell(
            &cell_xml(&format!(r#" t="{kind}""#), &format!("<v>{stored}</v>")),
            None,
        )?
        .ok_or("cell")?;
        assert_eq!(
            cell.formula,
            formula.map(|text| Formula::Text(text.to_owned())),
            "{kind} {stored}"
        );
        assert_eq!(cell.stale, formula.is_some());
    }
    Ok(())
}
