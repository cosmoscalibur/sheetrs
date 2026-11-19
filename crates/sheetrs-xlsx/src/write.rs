//! Writer of workbooks as XLSX packages.
//!
//! A workbook is written as `LibreOffice` writes it: numbers with 15 significant digits, text
//! in a shared string table and every sheet with its cells in row order. What the model does not
//! hold (styles, number formats, macros, external links) is not written.

mod number;
mod text;

use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;
use std::io::{Seek, Write};

use sheetrs_core::{
    CellPos, CellView, DateSystem, ErrorKind, FormulaGrammar, FormulaRef, OpaquePackage,
    PackageFormat, SheetOpaque, SheetView, ValueRef, Visibility, Workbook,
};
use zip::CompressionMethod;
use zip::write::{SimpleFileOptions, ZipWriter};

use crate::Error;
use number::number_text;
use text::push_text;

const MAIN_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const FLUSH_BYTES: usize = 1 << 16;

/// Writes `workbook` to `destination` as an XLSX package.
///
/// A workbook read from an XLSX package is written back into it: the entries the model does not
/// interpret, and the parts of the workbook and of each sheet around its cells, are copied as
/// they were read. Any other workbook gets the minimal parts of a package.
///
/// # Errors
///
/// [`Error::Unsupported`] when the workbook has formulas that are not in the grammar of XLSX or a
/// date system that XLSX has no flag for, when the date system differs from that of the package
/// it was read from, or when a sheet has no part in that package; [`Error::Zip`] or
/// [`Error::Io`] when the package cannot be written.
pub fn write<W: Write + Seek>(workbook: &Workbook, destination: W) -> Result<(), Error> {
    if matches!(workbook.date_system(), DateSystem::Custom { .. }) {
        return Err(Error::Unsupported(
            "XLSX counts dates from 1899-12-30 or from 1904-01-01".to_owned(),
        ));
    }
    if workbook.formula_grammar() != FormulaGrammar::Ooxml && !workbook.names().is_empty() {
        return Err(Error::Unsupported(
            "a name is not in the grammar of XLSX".to_owned(),
        ));
    }
    let package = workbook
        .opaque()
        .filter(|package| package.format == PackageFormat::Xlsx);
    if let Some(package) = package {
        if package.date_system != workbook.date_system() {
            return Err(Error::Unsupported(
                "the date system differs from that of the package the workbook was read from"
                    .to_owned(),
            ));
        }
        if workbook.sheets().any(|sheet| sheet.opaque().is_none()) {
            return Err(Error::Unsupported(
                "a sheet has no part in the package the workbook was read from".to_owned(),
            ));
        }
    }
    let mut archive = ZipWriter::new(destination);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    let mut strings = StringTable::default();
    let grammar = workbook.formula_grammar();
    if let Some(package) = package {
        // The content types come first, where readers look for them.
        let mut parts: Vec<_> = package.parts.iter().collect();
        parts.sort_by_key(|part| part.name != "[Content_Types].xml");
        for part in parts {
            archive.start_file(part.name.as_str(), options)?;
            archive.write_all(&part.bytes)?;
        }
        part(
            &mut archive,
            options,
            "xl/workbook.xml",
            &framed_workbook(workbook, package),
        )?;
        for sheet in workbook.sheets() {
            let Some(opaque) = sheet.opaque() else {
                continue;
            };
            archive.start_file(opaque.part.as_str(), options)?;
            let inline = package.shared_strings.is_none();
            write_sheet(
                &mut archive,
                sheet,
                grammar,
                &mut strings,
                Some((opaque, inline)),
            )?;
        }
        if let Some(path) = &package.shared_strings {
            part(&mut archive, options, path, &strings.part())?;
        }
    } else {
        let sheets = workbook.sheets().len();
        part(
            &mut archive,
            options,
            "[Content_Types].xml",
            &content_types(sheets),
        )?;
        part(&mut archive, options, "_rels/.rels", &root_relationships())?;
        part(
            &mut archive,
            options,
            "xl/workbook.xml",
            &workbook_part(workbook),
        )?;
        part(
            &mut archive,
            options,
            "xl/_rels/workbook.xml.rels",
            &workbook_relationships(sheets),
        )?;
        part(&mut archive, options, "xl/styles.xml", STYLES)?;
        for (index, sheet) in workbook.sheets().enumerate() {
            archive.start_file(format!("xl/worksheets/sheet{}.xml", index + 1), options)?;
            write_sheet(&mut archive, sheet, grammar, &mut strings, None)?;
        }
        part(
            &mut archive,
            options,
            "xl/sharedStrings.xml",
            &strings.part(),
        )?;
    }
    archive.finish()?;
    Ok(())
}

/// The workbook part of a package: its own text around the list of sheets, which comes from
/// the model.
fn framed_workbook(workbook: &Workbook, package: &OpaquePackage) -> String {
    let mut out = String::from_utf8_lossy(&package.frame.head).into_owned();
    out.push_str("<sheets>");
    for sheet in workbook.sheets() {
        let state = match sheet.visibility() {
            Visibility::Visible => "",
            // `LibreOffice` has no sheet that cannot be shown again: it writes it as hidden.
            Visibility::Hidden | Visibility::VeryHidden => r#" state="hidden""#,
        };
        out.push_str(r#"<sheet name=""#);
        push_text(&mut out, sheet.name());
        let _ = write!(out, r#""{state}"#);
        if let Some(opaque) = sheet.opaque().filter(|opaque| !opaque.entry.is_empty()) {
            out.push(' ');
            out.push_str(&opaque.entry);
        }
        out.push_str("/>");
    }
    out.push_str("</sheets>");
    let tail = String::from_utf8_lossy(&package.frame.tail);
    out.push_str(&with_defined_names(&tail, workbook, &package.defined_names));
    out
}

/// The tail of the workbook part with its `definedNames` element, which is added when there is
/// none, written from the names of the model and the `unread` elements of the package.
fn with_defined_names(tail: &str, workbook: &Workbook, unread: &[String]) -> String {
    let element = defined_names_element(workbook, unread);
    if let Some(start) = tail.find("<definedNames") {
        let end = if tail.get(start..).is_some_and(|rest| {
            rest.find('>')
                .is_some_and(|close| rest.as_bytes().get(close - 1) == Some(&b'/'))
        }) {
            tail.get(start..)
                .and_then(|rest| rest.find("/>"))
                .map(|at| start + at + 2)
        } else {
            tail.find("</definedNames>")
                .map(|at| at + "</definedNames>".len())
        };
        if let Some(end) = end {
            return format!(
                "{}{element}{}",
                tail.get(..start).unwrap_or_default(),
                tail.get(end..).unwrap_or_default()
            );
        }
        return tail.to_owned();
    }
    // The element goes after the external references and the function groups, where the schema
    // puts it.
    let at = ["</externalReferences>", "</functionGroups>"]
        .iter()
        .filter_map(|end| tail.find(end).map(|at| at + end.len()))
        .max()
        .unwrap_or(0);
    format!(
        "{}{element}{}",
        tail.get(..at).unwrap_or_default(),
        tail.get(at..).unwrap_or_default()
    )
}

/// The `definedNames` element, or nothing when there are no names.
fn defined_names_element(workbook: &Workbook, unread: &[String]) -> String {
    if workbook.names().is_empty() && unread.is_empty() {
        return String::new();
    }
    let mut out = String::from("<definedNames>");
    for name in workbook.names() {
        out.push_str(r#"<definedName name=""#);
        match name
            .name
            .strip_prefix("Excel_BuiltIn_")
            .filter(|rest| crate::names::BUILT_IN_NAMES.contains(rest))
        {
            Some(rest) => {
                out.push_str("_xlnm.");
                push_text(&mut out, rest);
            }
            None => push_text(&mut out, &name.name),
        }
        out.push('"');
        if let Some(scope) = name.scope {
            let _ = write!(out, r#" localSheetId="{scope}""#);
        }
        if name.hidden {
            out.push_str(r#" hidden="1""#);
        }
        out.push('>');
        push_text(&mut out, &name.expression);
        out.push_str("</definedName>");
    }
    for element in unread {
        out.push_str(element);
    }
    out.push_str("</definedNames>");
    out
}

fn part<W: Write + Seek>(
    archive: &mut ZipWriter<W>,
    options: SimpleFileOptions,
    name: &str,
    content: &str,
) -> Result<(), Error> {
    archive.start_file(name, options)?;
    archive.write_all(content.as_bytes())?;
    Ok(())
}

const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="10"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#;

fn content_types(sheets: usize) -> String {
    let mut out = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#,
    );
    for number in 1..=sheets {
        let _ = write!(
            out,
            r#"<Override PartName="/xl/worksheets/sheet{number}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        );
    }
    out.push_str("</Types>");
    out
}

fn root_relationships() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RELATIONSHIPS_NAMESPACE}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
    )
}

fn workbook_relationships(sheets: usize) -> String {
    let mut out = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for number in 1..=sheets {
        let _ = write!(
            out,
            r#"<Relationship Id="rId{number}" Type="{RELATIONSHIPS_NAMESPACE}/worksheet" Target="worksheets/sheet{number}.xml"/>"#
        );
    }
    let _ = write!(
        out,
        r#"<Relationship Id="rId{}" Type="{RELATIONSHIPS_NAMESPACE}/styles" Target="styles.xml"/><Relationship Id="rId{}" Type="{RELATIONSHIPS_NAMESPACE}/sharedStrings" Target="sharedStrings.xml"/></Relationships>"#,
        sheets + 1,
        sheets + 2
    );
    out
}

/// The workbook part: the date system, the sheet the workbook opens on, and the sheets.
fn workbook_part(workbook: &Workbook) -> String {
    let mut out = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{MAIN_NAMESPACE}" xmlns:r="{RELATIONSHIPS_NAMESPACE}">"#
    );
    if workbook.date_system() == DateSystem::Date1904 {
        out.push_str(r#"<workbookPr date1904="true"/>"#);
    }
    // The sheet that the workbook opens on is shown whatever its state says, so it is the first
    // that is shown.
    let active = workbook
        .sheets()
        .position(|sheet| sheet.visibility() == Visibility::Visible)
        .unwrap_or(0);
    let _ = write!(
        out,
        r#"<bookViews><workbookView activeTab="{active}"/></bookViews><sheets>"#
    );
    for (index, sheet) in workbook.sheets().enumerate() {
        let number = index + 1;
        let state = match sheet.visibility() {
            Visibility::Visible => "",
            // `LibreOffice` has no sheet that cannot be shown again: it writes it as hidden.
            Visibility::Hidden | Visibility::VeryHidden => r#" state="hidden""#,
        };
        out.push_str(r#"<sheet name=""#);
        push_text(&mut out, sheet.name());
        let _ = write!(out, r#"" sheetId="{number}"{state} r:id="rId{number}"/>"#);
    }
    out.push_str("</sheets>");
    out.push_str(&defined_names_element(workbook, &[]));
    out.push_str("</workbook>");
    out
}

/// The strings of the cells, in the order they were first written: text, and strings that
/// have formatting, which are kept as their package writes them.
#[derive(Default)]
struct StringTable {
    index: HashMap<(bool, String), usize>,
    order: Vec<(bool, String)>,
}

impl StringTable {
    fn number_of(&mut self, text: &str) -> usize {
        self.number(false, text)
    }

    fn number_of_formatted(&mut self, formatted: &str) -> usize {
        self.number(true, formatted)
    }

    fn number(&mut self, formatted: bool, text: &str) -> usize {
        let key = (formatted, text.to_owned());
        if let Some(found) = self.index.get(&key) {
            return *found;
        }
        let number = self.order.len();
        self.index.insert(key.clone(), number);
        self.order.push(key);
        number
    }

    fn part(&self) -> String {
        let mut out = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{MAIN_NAMESPACE}" count="{0}" uniqueCount="{0}">"#,
            self.order.len()
        );
        for (formatted, text) in &self.order {
            if *formatted {
                out.push_str(text);
            } else {
                out.push_str(r#"<si><t xml:space="preserve">"#);
                push_text(&mut out, text);
                out.push_str("</t></si>");
            }
        }
        out.push_str("</sst>");
        out
    }
}

/// The letters of column `col`, counted from zero.
fn column_letters(col: u32) -> String {
    let mut letters = Vec::new();
    let mut rest = col + 1;
    while rest > 0 {
        let digit = u8::try_from((rest - 1) % 26).unwrap_or(0);
        letters.push(b'A' + digit);
        rest = (rest - 1) / 26;
    }
    letters.reverse();
    String::from_utf8_lossy(&letters).into_owned()
}

fn address(pos: CellPos) -> String {
    format!("{}{}", column_letters(pos.col), pos.row + 1)
}

/// The shared groups of a sheet: the rectangle that holds the cells of each, as the first and
/// last cell, and the group each anchor cell carries the formula text of.
#[derive(Default)]
struct SharedGroups {
    extents: HashMap<u32, (CellPos, CellPos)>,
    anchors: HashMap<CellPos, u32>,
}

fn shared_groups(sheet: SheetView<'_>) -> SharedGroups {
    let mut extents: HashMap<u32, (CellPos, CellPos)> = HashMap::new();
    for (pos, cell) in sheet.cells() {
        let Some(FormulaRef::Shared { index }) = cell.formula else {
            continue;
        };
        let anchor = sheet.shared_anchor(index).unwrap_or(pos);
        let entry = extents.entry(index).or_insert((anchor, anchor));
        entry.0.row = entry.0.row.min(pos.row);
        entry.0.col = entry.0.col.min(pos.col);
        entry.1.row = entry.1.row.max(pos.row);
        entry.1.col = entry.1.col.max(pos.col);
    }
    let anchors = extents
        .keys()
        .filter_map(|index| Some((sheet.shared_anchor(*index)?, *index)))
        .collect();
    SharedGroups { extents, anchors }
}

/// How a sheet of a package is written: its own part around the cells, and whether text goes
/// inline because the package has no shared strings.
type Framing<'a> = Option<(&'a SheetOpaque, bool)>;

/// What writing the cells of one sheet needs to know about it.
struct CellContext<'a> {
    sheet: SheetView<'a>,
    shared: SharedGroups,
    /// The cells inside an array formula that hold the text it filled them with: the formula
    /// of the array covers them, so they are written without one.
    array_members: HashSet<CellPos>,
    grammar: FormulaGrammar,
    inline: bool,
}

/// Writes the part of `sheet`, with the text of its cells numbered in `strings`.
fn write_sheet<W: Write>(
    destination: &mut W,
    sheet: SheetView<'_>,
    grammar: FormulaGrammar,
    strings: &mut StringTable,
    framing: Framing<'_>,
) -> Result<(), Error> {
    let context = CellContext {
        sheet,
        shared: if sheet.has_shared_groups() {
            shared_groups(sheet)
        } else {
            SharedGroups::default()
        },
        array_members: sheet.array_members(),
        grammar,
        inline: framing.is_some_and(|(_, inline)| inline),
    };
    let mut out = sheet_head(sheet, framing);
    out.push_str("<sheetData>");
    let framed = framing.is_some();
    let mut cells = sheet.cells().peekable();
    let mut styled = sheet.styled_cells().peekable();
    let mut attributes = RowAttributes::new(if framed {
        sheet.row_attributes().collect()
    } else {
        Vec::new()
    });
    loop {
        let styled_row = if framed {
            styled.peek().map(|(pos, _)| pos.row)
        } else {
            None
        };
        let cell_row = cells.peek().map(|(pos, _)| pos.row);
        let Some(row) = [cell_row, styled_row, attributes.next_row()]
            .into_iter()
            .flatten()
            .min()
        else {
            break;
        };
        let _ = write!(out, r#"<row r="{}""#, row + 1);
        if let Some(text) = attributes.take(row) {
            out.push(' ');
            out.push_str(text);
        }
        out.push('>');
        loop {
            let next_cell = cells
                .peek()
                .filter(|(pos, _)| pos.row == row)
                .map(|(pos, _)| pos.col);
            let next_style = if framed {
                styled
                    .peek()
                    .filter(|(pos, _)| pos.row == row)
                    .map(|(pos, _)| pos.col)
            } else {
                None
            };
            let col = match (next_cell, next_style) {
                (None, None) => break,
                (Some(col), None) | (None, Some(col)) => col,
                (Some(a), Some(b)) => a.min(b),
            };
            let cell = if next_cell == Some(col) {
                cells.next().map(|(_, cell)| cell)
            } else {
                None
            };
            let style = if next_style == Some(col) {
                styled.next().map(|(_, style)| style)
            } else {
                None
            };
            write_cell(
                &mut out,
                &context,
                strings,
                CellPos { row, col },
                cell.as_ref(),
                style,
            )?;
            if out.len() >= FLUSH_BYTES {
                destination.write_all(out.as_bytes())?;
                out.clear();
            }
        }
        out.push_str("</row>");
    }
    out.push_str("</sheetData>");
    match framing {
        Some((opaque, _)) if !opaque.frame.tail.is_empty() => {
            out.push_str(&String::from_utf8_lossy(&opaque.frame.tail));
        }
        _ => out.push_str("</worksheet>"),
    }
    destination.write_all(out.as_bytes())?;
    Ok(())
}

/// The attributes of the rows of a sheet as runs, handed out row by row.
struct RowAttributes<'a> {
    runs: Vec<(u32, u32, &'a str)>,
    run: usize,
    next: Option<u32>,
}

impl<'a> RowAttributes<'a> {
    fn new(runs: Vec<(u32, u32, &'a str)>) -> Self {
        let next = runs.first().map(|(first, _, _)| *first);
        Self { runs, run: 0, next }
    }

    /// The first row that has attributes and was not handed out.
    fn next_row(&self) -> Option<u32> {
        self.next
    }

    /// The attributes of `row`, which are handed out once.
    fn take(&mut self, row: u32) -> Option<&'a str> {
        if self.next != Some(row) {
            return None;
        }
        let (_, last, text) = self.runs[self.run];
        if row < last {
            self.next = Some(row + 1);
        } else {
            self.run += 1;
            self.next = self.runs.get(self.run).map(|(first, _, _)| *first);
        }
        Some(text)
    }
}

/// The part of the sheet before its cells.
fn sheet_head(sheet: SheetView<'_>, framing: Framing<'_>) -> String {
    match framing {
        Some((opaque, _)) if !opaque.frame.head.is_empty() => {
            with_dimension(&String::from_utf8_lossy(&opaque.frame.head), sheet)
        }
        _ => format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN_NAMESPACE}">"#
        ),
    }
}

/// `head` with the `dimension` element it has, if any, set to the extent of the sheet.
fn with_dimension(head: &str, sheet: SheetView<'_>) -> String {
    let Some(start) = head.find("<dimension") else {
        return head.to_owned();
    };
    let Some(length) = head.get(start..).and_then(|rest| rest.find("/>")) else {
        return head.to_owned();
    };
    let reference = sheet.extent().map_or_else(
        || "A1".to_owned(),
        |(first, last)| {
            if first == last {
                address(first)
            } else {
                format!("{}:{}", address(first), address(last))
            }
        },
    );
    format!(
        r#"{}<dimension ref="{reference}"/>{}"#,
        head.get(..start).unwrap_or_default(),
        head.get(start + length + 2..).unwrap_or_default()
    )
}

/// Writes the cell at `pos`, which holds `cell` and has the style `style`; either may be absent.
fn write_cell(
    out: &mut String,
    context: &CellContext<'_>,
    strings: &mut StringTable,
    pos: CellPos,
    cell: Option<&CellView<'_>>,
    style: Option<u32>,
) -> Result<(), Error> {
    let style_attribute = style.map_or_else(String::new, |style| format!(r#" s="{style}""#));
    let Some(cell) = cell else {
        let _ = write!(out, r#"<c r="{}"{style_attribute}/>"#, address(pos));
        return Ok(());
    };
    let formula = if context.array_members.contains(&pos) {
        None
    } else {
        cell.formula
    };
    let kind = match (&cell.value, &formula) {
        (ValueRef::Number(_) | ValueRef::Empty, _) => "",
        (ValueRef::Bool(_), _) => r#" t="b""#,
        (ValueRef::Error(_), _) => r#" t="e""#,
        (ValueRef::Text(_), Some(_)) => r#" t="str""#,
        (ValueRef::Text(_), None) if context.inline => r#" t="inlineStr""#,
        (ValueRef::Text(_), None) => r#" t="s""#,
    };
    let _ = write!(out, r#"<c r="{}"{style_attribute}{kind}>"#, address(pos));
    if let Some(formula) = &formula {
        if context.grammar != FormulaGrammar::Ooxml {
            return Err(Error::Unsupported(
                "a formula is not in the grammar of XLSX".to_owned(),
            ));
        }
        write_formula(out, context.sheet, pos, formula, &context.shared);
    }
    match &cell.value {
        ValueRef::Empty => {}
        ValueRef::Number(number) => value(out, &number_text(*number)),
        ValueRef::Bool(flag) => value(out, if *flag { "1" } else { "0" }),
        ValueRef::Error(kind) => value(out, error_text(kind)),
        ValueRef::Text(text) if formula.is_some() => {
            out.push_str("<v>");
            push_text(out, text);
            out.push_str("</v>");
        }
        ValueRef::Text(text) if context.inline => {
            out.push_str(r#"<is><t xml:space="preserve">"#);
            push_text(out, text);
            out.push_str("</t></is>");
        }
        ValueRef::Text(text) => {
            // The formatting of a text stays while the text is as it was.
            let number = match context.sheet.rich_text(pos) {
                Some((plain, formatted)) if plain == *text => {
                    strings.number_of_formatted(formatted)
                }
                _ => strings.number_of(text),
            };
            value(out, &number.to_string());
        }
    }
    out.push_str("</c>");
    Ok(())
}

fn value(out: &mut String, text: &str) {
    out.push_str("<v>");
    out.push_str(text);
    out.push_str("</v>");
}

/// The error as XLSX writes it; the codes only `LibreOffice` has are written as `#VALUE!`.
fn error_text(kind: &ErrorKind) -> &str {
    match kind {
        ErrorKind::Other(_) => "#VALUE!",
        known => known.as_text(),
    }
}

fn write_formula(
    out: &mut String,
    sheet: SheetView<'_>,
    pos: CellPos,
    formula: &FormulaRef<'_>,
    extents: &SharedGroups,
) {
    match formula {
        FormulaRef::Shared { index } => {
            let _ = write!(out, r#"<f t="shared" si="{index}"/>"#);
        }
        FormulaRef::Text(text) => {
            let group = extents
                .anchors
                .get(&pos)
                .and_then(|index| extents.extents.get(index).map(|extent| (*index, extent)));
            let array = sheet
                .array_areas()
                .iter()
                .find(|(first, _)| *first == pos)
                .copied();
            if let Some((index, (first, last))) = group {
                let _ = write!(
                    out,
                    r#"<f t="shared" ref="{}:{}" si="{index}">"#,
                    address(*first),
                    address(*last)
                );
            } else if let Some((first, last)) = array {
                let _ = write!(
                    out,
                    r#"<f t="array" ref="{}:{}">"#,
                    address(first),
                    address(last)
                );
            } else {
                out.push_str("<f>");
            }
            push_text(out, text);
            out.push_str("</f>");
        }
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use sheetrs_core::{Cell, CellPos, Formula, Value};

    use super::*;

    fn constant(value: Value) -> Cell {
        Cell {
            value,
            formula: None,
            stale: false,
        }
    }

    fn pos(row: u32, col: u32) -> CellPos {
        CellPos { row, col }
    }

    #[test]
    fn a_written_workbook_reads_back_the_same() -> Result<(), Box<dyn std::error::Error>> {
        let mut workbook = Workbook::new();
        let mut sheet = workbook.sheet_builder("Datos & más");
        sheet.set(pos(0, 0), &constant(Value::Number(0.1 + 0.2)))?;
        sheet.set(pos(0, 1), &constant(Value::Text("a_x0041_\u{1}b".into())))?;
        sheet.set(pos(1, 0), &constant(Value::Bool(true)))?;
        sheet.set(pos(1, 1), &constant(Value::Error(ErrorKind::Na)))?;
        sheet.set(
            pos(2, 0),
            &Cell {
                value: Value::Number(3.0),
                formula: Some(Formula::Text("SUM(A1:A2)".to_owned())),
                stale: false,
            },
        )?;
        sheet.finish();
        let mut package = Cursor::new(Vec::new());
        write(&workbook, &mut package)?;
        package.set_position(0);
        let read = crate::read(package)?;
        let sheet = read.sheets().next().ok_or("no sheet")?;
        assert_eq!(sheet.name(), "Datos & más");
        let number = sheet.get(pos(0, 0)).ok_or("A1")?.value;
        assert_eq!(number, ValueRef::Number(0.3));
        assert_eq!(
            sheet.get(pos(0, 1)).ok_or("B1")?.value,
            ValueRef::Text("a_x0041_\u{1}b")
        );
        assert_eq!(
            sheet.get(pos(1, 0)).ok_or("A2")?.value,
            ValueRef::Bool(true)
        );
        assert_eq!(
            sheet.get(pos(1, 1)).ok_or("B2")?.value,
            ValueRef::Error(ErrorKind::Na)
        );
        let formula = sheet.get(pos(2, 0)).ok_or("A3")?;
        assert_eq!(formula.formula, Some(FormulaRef::Text("SUM(A1:A2)")));
        Ok(())
    }
}
