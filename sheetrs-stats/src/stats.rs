//! Workbook and per-sheet statistics computed from the format-agnostic model.

use serde::Serialize;
use sheetrs_core::reader::parser_utils::column_letters;
use sheetrs_core::reader::{Cell, CellSink, Sheet, Workbook};

/// Statistics for the whole workbook.
#[derive(Debug, Serialize)]
pub(crate) struct WorkbookStats {
    pub(crate) total_sheets: usize,
    pub(crate) total_named_ranges: usize,
    pub(crate) total_file_size: u64,
    pub(crate) total_used_cells: u64,
    pub(crate) total_formulas: u64,
    pub(crate) total_errors: u64,
    pub(crate) total_empty_cells: u64,
    pub(crate) total_comments: u64,
    pub(crate) total_charts: u64,
    pub(crate) total_tables: u64,
    pub(crate) total_xml_size: u64,
    pub(crate) sheets: Vec<SheetStats>,
}

/// Statistics for one sheet, in workbook order.
#[derive(Debug, Serialize)]
pub(crate) struct SheetStats {
    pub(crate) name: String,
    /// Cells holding a value or a formula.
    pub(crate) used_cells: u64,
    pub(crate) formulas: u64,
    /// Cells whose value (cached result for formulas) is an error.
    pub(crate) errors: u64,
    /// Cells without value or formula inside the used area, which spans from
    /// A1 to the last row and column holding a used cell.
    pub(crate) empty_cells: u64,
    /// Bottom right corner of the used area (`GV5439`), `None` for a sheet without used cells.
    pub(crate) end_of_sheet: Option<String>,
    /// Notes of the sheet's cells.
    pub(crate) comments: u64,
    pub(crate) charts: u64,
    /// Table objects of an XLSX, named database ranges of an ODS.
    pub(crate) tables: u64,
    /// Uncompressed size of the XML describing the sheet.
    pub(crate) xml_size: u64,
}

/// What is counted of the cells of one sheet as they are read.
#[derive(Debug, Default, Clone)]
pub(crate) struct Tally {
    used_cells: u64,
    formulas: u64,
    errors: u64,
    /// Bottom right corner of the used cells, `(row, col)` 0-based.
    used_area: Option<(u32, u32)>,
}

/// Counts the cells of each sheet as the reader hands them over, so the workbook's cells are never
/// held in memory.
#[derive(Default)]
pub(crate) struct StatsSink {
    sheets: Vec<Tally>,
}

impl CellSink for StatsSink {
    fn cell(&mut self, sheet: u16, (row, col): (u32, u32), cell: Cell) {
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
        tally.used_area = Some(match tally.used_area {
            Some((max_row, max_col)) => (max_row.max(row), max_col.max(col)),
            None => (row, col),
        });
    }
}

impl StatsSink {
    /// The tally of the sheet at `index`, empty for a sheet that gave no used cell.
    fn tally(&self, index: usize) -> Tally {
        self.sheets.get(index).cloned().unwrap_or_default()
    }
}

impl WorkbookStats {
    /// The statistics of `workbook`, whose cells were counted by `sink` while it was read.
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
            used_area,
        }: Tally,
    ) -> Self {
        let area = used_area.map_or(0, |(max_row, max_col)| {
            (u64::from(max_row) + 1) * (u64::from(max_col) + 1)
        });

        Self {
            name: sheet.name.clone(),
            used_cells,
            formulas,
            errors,
            empty_cells: area - used_cells,
            end_of_sheet: used_area
                .map(|(max_row, max_col)| format!("{}{}", column_letters(max_col), max_row + 1)),
            comments: sheet.comments as u64,
            charts: sheet.charts as u64,
            tables: sheet.tables as u64,
            xml_size: sheet.xml_size_bytes,
        }
    }
}

/// A cell counts as used when it holds a value or a formula; cells kept only
/// for their style do not.
fn is_used(cell: &Cell) -> bool {
    cell.is_formula() || !cell.value.is_empty()
}

/// Share of `part` in `total`, as a percentage (0 when `total` is 0).
pub(crate) fn percentage(part: u64, total: u64) -> f64 {
    if total == 0 {
        0.0
    } else {
        part as f64 / total as f64 * 100.0
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetrs_core::reader::{CellValue, read_workbook_from_reader_into};
    use sheetrs_core::violation::ExcelError;
    use std::io::Cursor;
    use std::sync::Arc;

    const TEST_XLSX: &[u8] = include_bytes!("../../tests/minimal_test.xlsx");
    const TEST_ODS: &[u8] = include_bytes!("../../tests/minimal_test.ods");

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

    fn stats_of_sheet(sheet: &Sheet, cells: Vec<((u32, u32), Cell)>) -> SheetStats {
        let mut sink = StatsSink::default();
        for (at, cell) in cells {
            sink.cell(0, at, cell);
        }
        SheetStats::from_sheet(sheet, sink.tally(0))
    }

    fn stats_for(bytes: &[u8], extension: &str) -> WorkbookStats {
        let mut sink = StatsSink::default();
        let workbook =
            read_workbook_from_reader_into(Cursor::new(bytes), Some(extension), &mut sink).unwrap();
        WorkbookStats::from_workbook(&workbook, &sink)
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
            let stats = stats_for(bytes, ext);

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

    #[test]
    fn xlsx_and_ods_cell_stats_match() {
        let xlsx = stats_for(TEST_XLSX, "xlsx");
        let ods = stats_for(TEST_ODS, "ods");

        assert_eq!(xlsx.total_sheets, ods.total_sheets);
        assert_eq!(xlsx.total_named_ranges, ods.total_named_ranges);
        for (x, o) in xlsx.sheets.iter().zip(&ods.sheets) {
            assert_eq!(x.name, o.name);
            assert_eq!(x.used_cells, o.used_cells, "used cells in {}", x.name);
            assert_eq!(x.formulas, o.formulas, "formulas in {}", x.name);
            assert_eq!(x.errors, o.errors, "errors in {}", x.name);
            assert_eq!(x.empty_cells, o.empty_cells, "empty cells in {}", x.name);
            assert_eq!(x.end_of_sheet, o.end_of_sheet, "end of {}", x.name);
            assert_eq!(x.comments, o.comments, "comments in {}", x.name);
        }
    }

    #[test]
    fn percentage_of_zero_total_is_zero() {
        assert_eq!(percentage(5, 0), 0.0);
        assert_eq!(percentage(1, 4), 25.0);
    }
}
