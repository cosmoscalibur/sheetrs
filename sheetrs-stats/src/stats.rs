//! Workbook and per-sheet statistics computed from the format-agnostic model.

use serde::Serialize;
use sheetrs_io::Book;
use sheetrs_io::Result;
use sheetrs_io::reader::parser_utils::column_letters;
use sheetrs_io::reader::{Cell, CellSink, Sheet, SheetState, Workbook, is_used};
use std::collections::BTreeMap;

/// Statistics for the whole workbook.
///
/// The serialized form (`sheetrs stats --format json`) is the totals and the `sheets`; `has_macros`
/// and `defined_names` are what the same read pass found out about the workbook, and are not in it.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize)]
pub struct WorkbookStats {
    pub total_sheets: usize,
    pub total_named_ranges: usize,
    pub total_file_size: u64,
    pub total_used_cells: u64,
    pub total_formulas: u64,
    pub total_errors: u64,
    pub total_empty_cells: u64,
    pub total_comments: u64,
    pub total_charts: u64,
    pub total_tables: u64,
    pub total_xml_size: u64,
    pub sheets: Vec<SheetStats>,
    /// Whether the workbook carries macros or VBA code (found, never run).
    #[serde(skip)]
    pub has_macros: bool,
    /// The defined names and what each refers to, as the file writes it, by name.
    #[serde(skip)]
    pub defined_names: BTreeMap<String, String>,
}

/// Statistics for one sheet, in workbook order.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize)]
pub struct SheetStats {
    pub name: String,
    /// Whether the sheet is shown, hidden or very hidden (not in the serialized form).
    #[serde(skip)]
    pub state: SheetState,
    /// Cells holding a value or a formula.
    pub used_cells: u64,
    pub formulas: u64,
    /// Cells whose value (cached result for formulas) is an error.
    pub errors: u64,
    /// Cells without value or formula inside the area from `A1` to the last used cell (the
    /// rectangle of the used range does not decide it: a sheet whose first used cell is `D10` has
    /// the cells of `A1:C9` empty).
    pub empty_cells: u64,
    /// Rectangle of the used cells (`D10:E11`), from the first used row and column to the last,
    /// as the used area of `LibreOffice`; `None` for a sheet without used cells. A single used cell
    /// gives `D10:D10`. It is [`Sheet::used_area`] of the model, the one the reader finds.
    pub used_range: Option<String>,
    /// Bottom right corner of the used area (`GV5439`), `None` for a sheet without used cells.
    pub end_of_sheet: Option<String>,
    /// Notes of the sheet's cells.
    pub comments: u64,
    pub charts: u64,
    /// Table objects of an XLSX, named database ranges of an ODS.
    pub tables: u64,
    /// Uncompressed size of the XML describing the sheet.
    pub xml_size: u64,
}

/// What is counted of the cells of one sheet as they are read. The used area is not here: it is the
/// reader's ([`Sheet::used_area`]), the one the linter works on.
#[derive(Debug, Default, Clone)]
pub(crate) struct Tally {
    used_cells: u64,
    formulas: u64,
    errors: u64,
}

/// Counts the cells of each sheet as the reader hands them over, so the workbook's cells are never
/// held in memory.
#[derive(Debug, Default)]
pub(crate) struct StatsSink {
    sheets: Vec<Tally>,
}

impl CellSink for StatsSink {
    fn cell(&mut self, sheet: u16, _at: (u32, u32), cell: Cell) {
        if !is_used(&cell) {
            return;
        }
        let sheet = usize::from(sheet);
        if self.sheets.len() <= sheet {
            self.sheets.resize(sheet + 1, Tally::default());
        }
        let tally = &mut self.sheets[sheet];
        tally.used_cells += 1;
        tally.formulas += u64::from(cell.is_formula());
        tally.errors += u64::from(cell.value.is_error());
    }
}

impl StatsSink {
    /// The tally of the sheet at `index`, empty for a sheet that gave no used cell.
    fn tally(&self, index: usize) -> Tally {
        self.sheets.get(index).cloned().unwrap_or_default()
    }
}

/// What a [`Book`] says about itself in numbers.
pub trait BookStats {
    /// The statistics of the workbook, counted in a single streaming pass: the cells are tallied as
    /// they are read and not kept, so the memory used does not depend on the size of the workbook.
    /// There is nothing to configure.
    ///
    /// ```no_run
    /// use sheetrs_io::Book;
    /// use sheetrs_stats::BookStats;
    ///
    /// # fn main() -> sheetrs_io::Result<()> {
    /// let stats = Book::from_file("book.xlsx")?.stats()?;
    /// println!("{} sheets, {} formulas", stats.total_sheets, stats.total_formulas);
    /// # Ok(()) }
    /// ```
    ///
    /// # Errors
    ///
    /// When the workbook cannot be read (see [`Book::workbook`]).
    fn stats(&self) -> Result<WorkbookStats>;
}

impl BookStats for Book {
    fn stats(&self) -> Result<WorkbookStats> {
        let mut sink = StatsSink::default();
        // The counts do not depend on the notation of a formula: the file's own (`workbook_into`)
        // skips the conversion of every ODS formula to Excel's that `analysis_into` does.
        let workbook = self.workbook_into(&mut sink)?;
        Ok(WorkbookStats::from_workbook(&workbook, &sink))
    }
}

impl WorkbookStats {
    /// The statistics of `workbook`, whose cells were counted by `sink` while it was read.
    #[must_use]
    pub(crate) fn from_workbook(workbook: &Workbook, sink: &StatsSink) -> Self {
        let sheets: Vec<SheetStats> = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(index, sheet)| SheetStats::from_sheet(sheet, sink.tally(index)))
            .collect();

        Self {
            total_sheets: sheets.len(),
            total_named_ranges: workbook.defined_names.len(),
            total_file_size: workbook.file_size_bytes,
            total_used_cells: sheets.iter().map(|s| s.used_cells).sum(),
            total_formulas: sheets.iter().map(|s| s.formulas).sum(),
            total_errors: sheets.iter().map(|s| s.errors).sum(),
            total_empty_cells: sheets.iter().map(|s| s.empty_cells).sum(),
            total_comments: sheets.iter().map(|s| s.comments).sum(),
            total_charts: sheets.iter().map(|s| s.charts).sum(),
            total_tables: sheets.iter().map(|s| s.tables).sum(),
            total_xml_size: sheets.iter().map(|s| s.xml_size).sum(),
            sheets,
            has_macros: workbook.has_macros,
            defined_names: workbook
                .defined_names
                .iter()
                .map(|(name, refers_to)| (name.clone(), refers_to.clone()))
                .collect(),
        }
    }
}

impl SheetStats {
    fn from_sheet(
        sheet: &Sheet,
        Tally {
            used_cells,
            formulas,
            errors,
        }: Tally,
    ) -> Self {
        let used_area = sheet.used_area;
        let corner = |row: u32, col: u32| format!("{}{}", column_letters(col), row + 1);

        Self {
            name: sheet.name.clone(),
            state: sheet.state,
            used_cells,
            formulas,
            errors,
            empty_cells: used_area.map_or(0, |area| area.cells_from_a1() - used_cells),
            used_range: used_area.map(|area| {
                format!(
                    "{}:{}",
                    corner(area.first_row, area.first_col),
                    corner(area.last_row, area.last_col)
                )
            }),
            end_of_sheet: used_area.map(|area| corner(area.last_row, area.last_col)),
            comments: sheet.comments as u64,
            charts: sheet.charts as u64,
            tables: sheet.tables as u64,
            xml_size: sheet.xml_size_bytes,
        }
    }
}

/// Share of `part` in `total`, as a percentage (0 when `total` is 0).
#[must_use]
pub fn percentage(part: u64, total: u64) -> f64 {
    if total == 0 {
        0.0
    } else {
        part as f64 / total as f64 * 100.0
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetrs_io::reader::ExcelError;
    use sheetrs_io::reader::{AreaSink, CellValue, UsedArea};
    use std::sync::Arc;

    const TEST_XLSX: &[u8] = include_bytes!("../../tests/lint_book.xlsx");
    const TEST_ODS: &[u8] = include_bytes!("../../tests/lint_book.ods");

    fn cell(row: u32, col: u32, value: CellValue, formula: Option<&str>) -> ((u32, u32), Cell) {
        (
            (row, col),
            Cell {
                value,
                formula: formula.map(Into::into),
                ..Default::default()
            },
        )
    }

    /// The statistics of a sheet `Data` that the reader gave `cells`.
    fn stats_of(cells: Vec<((u32, u32), Cell)>) -> SheetStats {
        stats_of_sheet(&Sheet::new("Data".to_string(), 0), cells)
    }

    /// The cells go through the reader's accumulator of the used area, as they do when a file is read.
    fn stats_of_sheet(sheet: &Sheet, cells: Vec<((u32, u32), Cell)>) -> SheetStats {
        let mut sink = StatsSink::default();
        let mut areas = AreaSink::new(&mut sink);
        for (at, cell) in cells {
            areas.cell(0, at, cell);
        }
        let mut sheet = sheet.clone();
        sheet.used_area = areas.into_areas().first().and_then(|a| a.used);
        SheetStats::from_sheet(&sheet, sink.tally(0))
    }

    fn stats_for(bytes: &[u8]) -> WorkbookStats {
        Book::from_bytes(bytes).unwrap().stats().unwrap()
    }

    #[test]
    fn counts_used_formula_error_and_empty_cells() {
        // B3 bounds a 3x2 used area holding 4 used cells.
        let stats = stats_of(vec![
            cell(0, 0, CellValue::Number(1.0), None),
            cell(0, 1, CellValue::Text(Arc::from("x")), None),
            cell(1, 0, CellValue::Number(2.0), Some("A1*2")),
            cell(2, 1, CellValue::Error(ExcelError::DivZero), Some("A1/0")),
        ]);

        assert_eq!(stats.used_cells, 4);
        assert_eq!(stats.formulas, 2);
        assert_eq!(stats.errors, 1);
        assert_eq!(stats.empty_cells, 2);
        assert_eq!(stats.end_of_sheet.as_deref(), Some("B3"));
    }

    #[test]
    fn the_used_range_is_the_rectangle_of_the_used_cells() {
        // D10 and E11 (0-based (9, 3) and (10, 4)), given out of order.
        let stats = stats_of(vec![
            cell(10, 4, CellValue::Text(Arc::from("y")), None),
            cell(9, 3, CellValue::Text(Arc::from("x")), None),
            cell(0, 0, CellValue::Empty, None),
        ]);
        assert_eq!(stats.used_range.as_deref(), Some("D10:E11"));
        assert_eq!(stats.end_of_sheet.as_deref(), Some("E11"));
        assert_eq!(stats.used_cells, 2);
        // Counted from A1: A1:E11 is 55 cells, 2 of them used.
        assert_eq!(stats.empty_cells, 53);
        assert_eq!(stats_of(vec![]).used_range, None);
    }

    /// The used cells D10 and E11 of a workbook written in each format, read from bytes.
    #[test]
    fn the_used_range_of_a_file_starts_at_its_first_used_cell() {
        let xlsx = {
            use std::io::Write;
            let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            let opts = zip::write::SimpleFileOptions::default();
            for (name, body) in [
                (
                    "[Content_Types].xml",
                    r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
                ),
                (
                    "_rels/.rels",
                    r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
                ),
                (
                    "xl/workbook.xml",
                    r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
                ),
                (
                    "xl/_rels/workbook.xml.rels",
                    r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
                ),
                (
                    "xl/worksheets/sheet1.xml",
                    r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E11"/><sheetData><row r="10"><c r="D10" t="inlineStr"><is><t>x</t></is></c></row><row r="11"><c r="E11"><v>2</v></c></row></sheetData></worksheet>"#,
                ),
            ] {
                zip.start_file(name, opts).unwrap();
                zip.write_all(body.as_bytes()).unwrap();
            }
            zip.finish().unwrap().into_inner()
        };
        let ods = {
            use std::io::Write;
            let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            let stored = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", stored).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")
                .unwrap();
            zip.start_file("content.xml", zip::write::SimpleFileOptions::default())
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.2"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-column table:number-columns-repeated="3"/><table:table-column/><table:table-column/><table:table-row table:number-rows-repeated="9"><table:table-cell table:number-columns-repeated="5"/></table:table-row><table:table-row><table:table-cell table:number-columns-repeated="3"/><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell table:number-columns-repeated="4"/><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#).unwrap();
            zip.finish().unwrap().into_inner()
        };
        for (bytes, ext) in [(xlsx, "xlsx"), (ods, "ods")] {
            let stats = stats_for(&bytes);
            let sheet = &stats.sheets[0];
            assert_eq!(sheet.used_range.as_deref(), Some("D10:E11"), "{ext}");
            assert_eq!(sheet.used_cells, 2, "{ext}");
            assert_eq!(sheet.empty_cells, 53, "{ext}");
        }
    }

    /// An XLSX with one sheet `Sheet1` whose `sheetData` is `rows` (and a `dimension` that lies).
    fn xlsx_with(rows: &str) -> Vec<u8> {
        use std::io::Write;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let opts = zip::write::SimpleFileOptions::default();
        let sheet = format!(
            r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:ZZ500"/><sheetData>{rows}</sheetData></worksheet>"#
        );
        for (name, body) in [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", sheet.as_str()),
        ] {
            zip.start_file(name, opts).unwrap();
            zip.write_all(body.as_bytes()).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    /// An ODS with one sheet `Sheet1` whose table holds `rows`.
    fn ods_with(rows: &str) -> Vec<u8> {
        use std::io::Write;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")
            .unwrap();
        zip.start_file("content.xml", zip::write::SimpleFileOptions::default())
            .unwrap();
        let content = format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.2"><office:automatic-styles><style:style style:name="ce1" style:family="table-cell"/></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1">{rows}</table:table></office:spreadsheet></office:body></office:document-content>"#
        );
        zip.write_all(content.as_bytes()).unwrap();
        zip.finish().unwrap().into_inner()
    }

    /// The used area is one: the model (what the linter walks) and the statistics (one streaming
    /// pass) give the same bounds for the same file, in both formats, and neither is moved by a cell
    /// kept only for its style, by a hidden row or column, or by the `dimension` of an XLSX.
    ///
    /// The sheet: values at AA5 and AC7 (the first used column is far to the right), a cell that is
    /// only styled (near `A1` and at AF9), and a hidden row 20.
    #[test]
    fn the_model_and_the_statistics_give_the_same_used_area() {
        let xlsx = xlsx_with(
            r#"<row r="2"><c r="B2" s="1"/></row><row r="5"><c r="AA5" t="inlineStr"><is><t>x</t></is></c></row><row r="7"><c r="AC7"><f>1+1</f><v>2</v></c></row><row r="9"><c r="AF9" s="1"/></row><row r="20" hidden="1"/>"#,
        );
        // The same, with a styled B1 for the first one.
        let ods = ods_with(
            r#"<table:table-column table:number-columns-repeated="60"/><table:table-row><table:table-cell table:number-columns-repeated="1"/><table:table-cell table:style-name="ce1"/></table:table-row><table:table-row table:number-rows-repeated="3"><table:table-cell/></table:table-row><table:table-row><table:table-cell table:number-columns-repeated="26"/><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell/></table:table-row><table:table-row><table:table-cell table:number-columns-repeated="28"/><table:table-cell table:formula="of:=1+1" office:value-type="float" office:value="2"/></table:table-row><table:table-row><table:table-cell/></table:table-row><table:table-row><table:table-cell table:number-columns-repeated="31"/><table:table-cell table:style-name="ce1"/></table:table-row><table:table-row table:number-rows-repeated="10"><table:table-cell/></table:table-row><table:table-row table:visibility="collapse"><table:table-cell/></table:table-row>"#,
        );
        for (bytes, ext) in [(xlsx, "xlsx"), (ods, "ods")] {
            let book = Book::from_bytes(bytes).unwrap();
            let expected = UsedArea {
                first_row: 4,
                first_col: 26,
                last_row: 6,
                last_col: 28,
            };
            // The model, as the linter reads it, with the cells in memory.
            let model = book.workbook().unwrap();
            assert_eq!(model.sheets[0].used_area, Some(expected), "{ext} model");
            // The A1 extent of the linter comes from the same bounds.
            // What the file keeps beyond the content (the styled AF9, the hidden row 20) is the other
            // rectangle, found by the same accumulator: the linter's blank rows and columns work on it.
            assert_eq!(
                model.sheets[0].kept_extent(),
                Some((20, 32)),
                "{ext} extent"
            );
            // The streaming pass, which keeps no cells.
            let stats = book.stats().unwrap();
            let sheet = &stats.sheets[0];
            assert_eq!(sheet.used_range.as_deref(), Some("AA5:AC7"), "{ext} stats");
            assert_eq!(sheet.end_of_sheet.as_deref(), Some("AC7"), "{ext}");
            assert_eq!(sheet.used_cells, 2, "{ext}");
            // From A1: 7 rows by 29 columns, 2 of them used.
            assert_eq!(sheet.empty_cells, 7 * 29 - 2, "{ext}");
            // The same file through the analysis form (the linter's).
            let analysis = book.analysis_into(&mut NoCells).unwrap();
            assert_eq!(
                analysis.sheets[0].used_area,
                Some(expected),
                "{ext} analysis"
            );
        }
    }

    /// A sink that drops the cells: the area must come out all the same.
    struct NoCells;

    impl CellSink for NoCells {
        fn cell(&mut self, _sheet: u16, _at: (u32, u32), _cell: Cell) {}
    }

    #[test]
    fn the_end_of_a_sheet_is_the_corner_of_the_used_area() {
        let stats = stats_of(vec![cell(5438, 203, CellValue::Number(1.0), None)]);
        assert_eq!(stats.end_of_sheet.as_deref(), Some("GV5439"));
        assert_eq!(stats_of(vec![]).end_of_sheet, None);
    }

    #[test]
    fn notes_charts_and_tables_come_from_the_sheet() {
        let mut sheet = Sheet::new("Data".to_string(), 0);
        (sheet.comments, sheet.charts, sheet.tables) = (3, 2, 1);
        let stats = stats_of_sheet(&sheet, vec![]);
        assert_eq!((stats.comments, stats.charts, stats.tables), (3, 2, 1));
    }

    #[test]
    fn style_only_cells_are_neither_used_nor_extend_the_area() {
        let stats = stats_of(vec![
            cell(0, 0, CellValue::Number(1.0), None),
            cell(9, 9, CellValue::Empty, None),
        ]);

        assert_eq!(stats.used_cells, 1);
        assert_eq!(stats.empty_cells, 0);
    }

    #[test]
    fn empty_sheet_has_no_cells() {
        let stats = stats_of(vec![]);

        assert_eq!(stats.used_cells, 0);
        assert_eq!(stats.empty_cells, 0);
    }

    #[test]
    fn workbook_totals_add_up_sheet_stats() {
        for (bytes, ext) in [(TEST_XLSX, "xlsx"), (TEST_ODS, "ods")] {
            let stats = stats_for(bytes);

            assert_eq!(stats.total_sheets, stats.sheets.len(), "{ext}");
            let sum = |f: fn(&SheetStats) -> u64| stats.sheets.iter().map(f).sum::<u64>();
            assert_eq!(stats.total_used_cells, sum(|s| s.used_cells), "{ext}");
            assert_eq!(stats.total_formulas, sum(|s| s.formulas), "{ext}");
            assert_eq!(stats.total_errors, sum(|s| s.errors), "{ext}");
            assert_eq!(stats.total_empty_cells, sum(|s| s.empty_cells), "{ext}");
            assert_eq!(stats.total_xml_size, sum(|s| s.xml_size), "{ext}");
            assert!(stats.sheets.iter().all(|s| s.xml_size > 0), "{ext}");
        }
    }

    /// The cells, formulas and errors of each sheet of each file are those `LibreOffice` reports for
    /// that file (`sheetrs-io/tests/assets/oracle/lint_book.<ext>.oracle.json`, made by
    /// `make_oracle.py`).
    ///
    /// The differences that nobody approved are listed in `PENDING`, and the test fails when one of
    /// them goes. There is none.
    #[test]
    fn cell_stats_are_what_libreoffice_reports_for_each_file() {
        // (file, sheet, what is counted, what the reader counts, what LibreOffice counts)
        const PENDING: &[(&str, &str, &str, u64, u64)] = &[];
        let mut found: Vec<(String, String, &str, u64, u64)> = Vec::new();
        for (bytes, ext) in [(TEST_XLSX, "xlsx"), (TEST_ODS, "ods")] {
            let path = format!(
                "{}/../sheetrs-io/tests/assets/oracle/lint_book.{ext}.oracle.json",
                env!("CARGO_MANIFEST_DIR")
            );
            let oracle: serde_json::Value =
                serde_json::from_slice(&std::fs::read(path).unwrap()).unwrap();
            let stats = stats_for(bytes);
            let libre_sheets = oracle["sheets"].as_array().unwrap();
            assert_eq!(stats.total_sheets, libre_sheets.len(), "{ext}");
            for sheet in &stats.sheets {
                let cells: Vec<&serde_json::Value> = oracle["cells"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter(|c| c["sheet"] == sheet.name.as_str())
                    .collect();
                let count = |keep: fn(&serde_json::Value) -> bool| {
                    cells.iter().filter(|c| keep(c)).count() as u64
                };
                for (what, ours, libre) in [
                    ("used cells", sheet.used_cells, cells.len() as u64),
                    (
                        "formulas",
                        sheet.formulas,
                        count(|c| c["type"] == "FORMULA"),
                    ),
                    ("errors", sheet.errors, count(|c| c["error"] != 0)),
                ] {
                    if ours != libre {
                        found.push((ext.to_string(), sheet.name.clone(), what, ours, libre));
                    }
                }
            }
        }
        let pending: Vec<(String, String, &str, u64, u64)> = PENDING
            .iter()
            .map(|&(f, s, w, o, l)| (f.to_string(), s.to_string(), w, o, l))
            .collect();
        assert_eq!(found, pending);
    }

    /// What `tests/make_lint_book.py` puts in the workbook, counted by hand: twenty sheets (one of them
    /// empty, two hidden); the names `Rate`, `TaxRate`, `Local` (defined for the book and for `Calc`,
    /// counted once) and `Gone`; one note (`Inputs!B1`); one chart (`Summary`); nine error cells
    /// (`Errors!A1:A5`, the two of the loop in `Loops`, `Summary!B5`, which reads both, and
    /// `Vulnerable!E4`, a `MODE` of values without repeats); no table. The two files agree with it, each
    /// on its own account.
    #[test]
    fn the_statistics_of_the_lint_book_are_what_the_workbook_holds() {
        for (bytes, ext) in [(TEST_XLSX, "xlsx"), (TEST_ODS, "ods")] {
            let stats = stats_for(bytes);
            assert_eq!(stats.total_sheets, 20, "{ext}");
            assert_eq!(stats.total_named_ranges, 4, "{ext}");
            assert_eq!(stats.total_comments, 1, "{ext}");
            assert_eq!(stats.total_charts, 1, "{ext}");
            assert_eq!(stats.total_tables, 0, "{ext}");
            assert_eq!(stats.total_errors, 9, "{ext}");
            let sheet = |name: &str| stats.sheets.iter().find(|s| s.name == name).unwrap();
            assert_eq!(sheet("Errors").errors, 5, "{ext}");
            assert_eq!(sheet("Loops").errors, 2, "{ext}");
            assert_eq!(sheet("Empty").used_cells, 0, "{ext}");
            assert_eq!(sheet("Inputs").comments, 1, "{ext}");
            assert_eq!(sheet("Summary").charts, 1, "{ext}");
            assert_eq!(sheet("Seq").end_of_sheet.as_deref(), Some("D19"), "{ext}");
        }
    }

    #[test]
    fn percentage_of_zero_total_is_zero() {
        assert_eq!(percentage(5, 0), 0.0);
        assert_eq!(percentage(1, 4), 25.0);
    }

    #[test]
    fn stats_gives_the_same_sheets_from_a_file_and_from_bytes() {
        let tests = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../tests");
        for (name, bytes) in [("lint_book.xlsx", TEST_XLSX), ("lint_book.ods", TEST_ODS)] {
            let file = Book::from_file(tests.join(name)).unwrap();
            let by_file = file.stats().unwrap();
            let by_bytes = Book::from_bytes(bytes).unwrap().stats().unwrap();
            assert_eq!(
                serde_json::to_value(&by_file.sheets).unwrap(),
                serde_json::to_value(&by_bytes.sheets).unwrap(),
                "{name}"
            );
            assert_eq!(by_file.total_formulas, by_bytes.total_formulas);
            assert_eq!(by_bytes.total_file_size, 0, "bytes carry no size");
            assert!(by_file.total_file_size > 0);
        }
    }

    #[test]
    fn stats_reports_a_workbook_that_cannot_be_read_with_a_typed_error() {
        // What is not a workbook, or is not there, is known when the source is made.
        assert!(matches!(
            Book::from_bytes(&b"not a zip"[..]),
            Err(sheetrs_io::Error::Unsupported(_))
        ));
        assert!(matches!(
            Book::from_file("tests/does-not-exist.xlsx"),
            Err(sheetrs_io::Error::Open { .. })
        ));
    }
}
