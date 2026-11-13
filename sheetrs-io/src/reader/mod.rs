//! Excel/ODS file reader using custom XML parsers

use crate::error::{Error, Result};

use sheetrs_formula::DateSystem;
use std::collections::HashMap;
use std::path::Path;
use zip::ZipArchive;

mod detect;
pub mod excel_error;
mod external;
mod format;
pub mod ods_parser;
mod ods_styles;
pub mod parser_utils;
mod selection;
pub mod workbook;
pub mod xlsx_parser;

use self::ods_parser::OdsReader;
use self::xlsx_parser::XlsxReader;
pub use excel_error::ExcelError;
pub use external::{ExternalUrl, external_url};
pub use format::Format;
pub use selection::{Properties, Selection, SheetInfo, Window};
pub use sheetrs_formula::grid::Grid;
pub use workbook::{
    Cell, CellValue, DocumentSettings, ExternalWorkbook, IterationSettings, NameConflict,
    NameScope, NumFmt, Sheet, SheetState, Workbook,
};

/// Trait for spreadsheet format readers
///
/// Where a reader puts the cells it reads. The default readers collect them into each sheet's
/// `cells` map; a consumer with its own storage (the compiler's model) implements this to receive
/// them as they are read, with no intermediate copy.
pub trait CellSink {
    /// An upper bound on the cells of the sheet about to be read: the area of its `dimension`.
    fn hint(&mut self, _sheet: u16, _cells: usize) {}

    /// The cell of the sheet at `(row, column)`, 0-based. A later cell at the same position
    /// replaces an earlier one.
    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell);

    /// The text of the note of the cell at `(row, column)` (0-based), when the selection asks for
    /// it ([`Properties::comment`]): the paragraphs of the note joined by `\n`, the runs of a rich
    /// text one after the other. It comes **before** the cell it belongs to (which may not come at
    /// all: a note can sit on a cell with nothing else).
    fn comment(&mut self, _sheet: u16, _at: (u32, u32), _text: &str) {}

    /// The sheet `index` called `name` begins: its cells follow. Called once per sheet read, before
    /// its first cell (and its hint).
    fn sheet(&mut self, _index: u16, _name: &str) {}

    /// The rows `first..=last` (0-based) of the sheet are hidden **by a filter** (an autofilter),
    /// as opposed to hidden by hand. Called only when the selection asks for it
    /// ([`Selection::with_filtered_rows`](crate::reader::Selection::with_filtered_rows)). An ODS says
    /// it row by row (`table:visibility="filter"`), before the cells of the row; an XLSX only after
    /// the whole sheet is read, because the range of the autofilter comes after the rows: its hidden
    /// rows inside that range are the filtered ones. Whoever uses it keeps the ranges and applies
    /// them when the sheet is done, in either format.
    fn filtered_rows(&mut self, _sheet: u16, _first: u32, _last: u32) {}

    /// The cell of the sheet at `(row, column)` (0-based) is a member of an array formula, and sits
    /// `offset` = (rows, columns) below and to the right of the top-left cell of the block the
    /// formula fills. It comes **before** the [`cell`](CellSink::cell) it is about, and only for the
    /// members that are not the top-left cell (a formula that holds one value has none): each member
    /// shows its own element of the result.
    fn array_member(&mut self, _sheet: u16, _at: (u32, u32), _offset: (u32, u32)) {}
}

/// A sink that hands on the cells of `inner` with their formulas in Excel's notation, the
/// representation of the analysis tools (see [`ods_parser::ods_formula_to_excel`]): an XLSX already
/// keeps them so, an ODS keeps `OpenFormula`. Whoever needs a file's own notation reads without it.
pub struct ExcelFormulas<'a> {
    inner: &'a mut dyn CellSink,
    ods: bool,
    names: Vec<String>,
}

impl<'a> ExcelFormulas<'a> {
    /// `format` is the format of the file being read.
    pub fn new(inner: &'a mut dyn CellSink, format: Format) -> Self {
        Self {
            inner,
            ods: format == Format::Ods,
            names: Vec::new(),
        }
    }
}

impl CellSink for ExcelFormulas<'_> {
    fn hint(&mut self, sheet: u16, cells: usize) {
        self.inner.hint(sheet, cells);
    }

    fn sheet(&mut self, index: u16, name: &str) {
        let i = usize::from(index);
        if self.names.len() <= i {
            self.names.resize(i + 1, String::new());
        }
        self.names[i] = name.to_string();
        self.inner.sheet(index, name);
    }

    fn filtered_rows(&mut self, sheet: u16, first: u32, last: u32) {
        self.inner.filtered_rows(sheet, first, last);
    }

    fn array_member(&mut self, sheet: u16, at: (u32, u32), offset: (u32, u32)) {
        self.inner.array_member(sheet, at, offset);
    }

    fn cell(&mut self, sheet: u16, at: (u32, u32), mut cell: Cell) {
        if self.ods
            && let Some(native) = cell.formula.as_deref()
        {
            let home = self
                .names
                .get(usize::from(sheet))
                .map_or("", String::as_str);
            cell.formula = Some(ods_parser::ods_formula_to_excel(native, home).into());
        }
        self.inner.cell(sheet, at, cell);
    }
}

/// The sink the ordinary readers use: one [`Grid`] per sheet.
#[derive(Default)]
pub(crate) struct GridSink {
    pub sheets: Vec<Grid<Cell>>,
}

impl CellSink for GridSink {
    fn hint(&mut self, sheet: u16, cells: usize) {
        self.grid(sheet).reserve(cells.min(HINT_LIMIT));
    }

    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell) {
        self.grid(sheet).insert(at, cell);
    }
}

/// A `dimension` can be much larger than the cells a sheet holds: reserve at most this many.
const HINT_LIMIT: usize = 1 << 20;

impl GridSink {
    fn grid(&mut self, sheet: u16) -> &mut Grid<Cell> {
        let i = usize::from(sheet);
        if self.sheets.len() <= i {
            self.sheets.resize_with(i + 1, Grid::new);
        }
        &mut self.sheets[i]
    }

    /// The cells read go to the sheets, with no room left over.
    pub(crate) fn into_sheets(self, sheets: &mut [Sheet]) {
        for (sheet, mut cells) in sheets.iter_mut().zip(self.sheets) {
            cells.shrink_to_fit();
            sheet.cells = cells;
        }
    }
}

/// The names that were defined in more than one scope, from every scope each one was defined in:
/// alphabetical order of the name, the workbook scope first and then the sheets in the order they
/// were found.
pub(crate) fn conflicts_of(scopes: HashMap<String, Vec<NameScope>>) -> Vec<NameConflict> {
    let mut conflicts: Vec<NameConflict> = scopes
        .into_iter()
        .filter(|(_, scopes)| scopes.len() > 1)
        .map(|(name, mut scopes)| {
            // Stable: the workbook (`false`) before the sheets, which keep their order.
            scopes.sort_by_key(|scope| matches!(scope, NameScope::Sheet(_)));
            NameConflict { name, scopes }
        })
        .collect();
    conflicts.sort_by(|a, b| a.name.cmp(&b.name));
    conflicts
}

/// This trait defines the common interface that different file format parsers (XLSX, ODS)
/// must implement. It allows the main logic to abstract over specific file formats.
pub trait WorkbookReader {
    /// Read the sheets, handing their cells to `sink`; the returned sheets have empty `cells`.
    /// The one place a format's cells are read: every consumer gets them through a sink.
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_sheets_into(&mut self, sink: &mut dyn CellSink) -> Result<Vec<Sheet>>;

    /// Read what a [`Selection`] asks for and nothing else, handing the cells to `sink`: the
    /// sheets that are not selected are not parsed, the reading of a sheet stops after the last
    /// row a window reaches, and the properties not asked for are not built. Returns the sheets
    /// the reader went through, in the order of the workbook (an ODS ends with the last sheet
    /// selected). A cell reaches the sink with the index of its sheet in that order.
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_selection_into(
        &mut self,
        selection: &Selection,
        sink: &mut dyn CellSink,
    ) -> Result<Vec<SheetInfo>>;

    /// Read all sheets from the workbook, with their cells in a [`Grid`].
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_sheets(&mut self) -> Result<Vec<Sheet>> {
        let mut sink = GridSink::default();
        let mut sheets = self.read_sheets_into(&mut sink)?;
        sink.into_sheets(&mut sheets);
        Ok(sheets)
    }

    /// Read defined names (named ranges) from the workbook
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_defined_names(&mut self) -> Result<HashMap<String, String>>;

    /// The defined names that the document defines in more than one scope (see
    /// [`NameConflict`]), in alphabetical order. Reserved names (`_xlnm.`) are not names of the
    /// user and are left out.
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_name_conflicts(&mut self) -> Result<Vec<NameConflict>>;

    /// Read the list of hidden sheet names
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_hidden_sheets(&mut self) -> Result<Vec<String>>;

    /// Check if the workbook contains macros (VBA, Basic, Scripts)
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn has_macros(&mut self) -> Result<bool>;

    /// Read external file links (generic links)
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_external_links(&mut self) -> Result<Vec<String>>;

    /// Read external workbook references with their indices
    ///
    /// The indices correspond to the reference ID used in formulas (e.g., `[1]Sheet1!A1`).
    /// For XLSX, this maps to the externalLink relations.
    /// For ODS, this maps to the order of appearance or explicit file references.
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_external_workbooks(&mut self) -> Result<Vec<ExternalWorkbook>>;

    /// Read the last modification date (parsed from file metadata)
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_modified_date(&mut self) -> Result<Option<chrono::DateTime<chrono::Utc>>>;

    /// The date system of the workbook: the date of serial 0.
    ///
    /// # Errors
    ///
    /// [`Error::UnsupportedDateSystem`] when an ODS declares a null date that is not one of the three
    /// date systems of `LibreOffice` Calc; otherwise when the part cannot be read or is malformed.
    fn read_date_system(&mut self) -> Result<DateSystem>;

    /// Read the document properties the file declares (title, author, dates…), as stored.
    ///
    /// The names are the same in XLSX and ODS where both have the property: `title`, `subject`,
    /// `description`, `keywords`, `creator` (who made it), `last_modified_by`, `created`,
    /// `modified`, `generator` (the application); the others keep the name the file gives them.
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_properties(&mut self) -> Result<Vec<(String, String)>>;

    /// Read the calculation settings stored in the document, as stored.
    ///
    /// # Errors
    ///
    /// When the part cannot be read or is malformed.
    fn read_settings(&mut self) -> Result<DocumentSettings>;
}

/// Read a workbook from a generic reader (must implement Read + Seek)
///
/// # Arguments
///
/// * `reader` - Any source that implements `std::io::Read` and `std::io::Seek` (e.g., `File`, `Cursor`);
///   its format is read from its content (see [`Format::detect_reader`])
///
/// # Errors
///
/// When the file cannot be read or is not a workbook of a supported format.
pub(crate) fn read_workbook_from_reader<R: std::io::Read + std::io::Seek>(
    reader: R,
) -> Result<Workbook> {
    let mut sink = GridSink::default();
    let mut workbook = read_workbook_inner(reader, &mut sink, true)?;
    sink.into_sheets(&mut workbook.sheets);
    Ok(workbook)
}

/// [`read_workbook_from_reader`] with the cells handed to `sink`.
///
/// # Errors
///
/// When the file cannot be read or is not a workbook of a supported format.
pub(crate) fn read_workbook_from_reader_into<R: std::io::Read + std::io::Seek>(
    reader: R,
    sink: &mut dyn CellSink,
) -> Result<Workbook> {
    read_workbook_inner(reader, sink, false)
}

/// [`read_workbook_from_reader_into`] with the formulas in the analysis form: Excel's notation (see
/// [`ExcelFormulas`]), which is what the linter, the statistics and the compiler work on.
///
/// # Errors
///
/// When the file cannot be read or is not a workbook of a supported format.
pub(crate) fn read_analysis_from_reader_into<R: std::io::Read + std::io::Seek>(
    reader: R,
    sink: &mut dyn CellSink,
) -> Result<Workbook> {
    read_workbook_inner(reader, sink, true)
}

fn read_workbook_inner<R: std::io::Read + std::io::Seek>(
    reader: R,
    sink: &mut dyn CellSink,
    excel_formulas: bool,
) -> Result<Workbook> {
    let mut archive = open_archive(reader)?;
    let format = detect::detect_archive(&mut archive)?;

    let (
        sheets,
        defined_names,
        name_conflicts,
        hidden_sheets,
        has_macros,
        external_workbooks,
        modified_date,
        date_system,
        settings,
        properties,
    ) = if format == Format::Ods {
        let mut reader = OdsReader::new(&mut archive)?;
        let sheets = if excel_formulas {
            reader.read_sheets_into(&mut ExcelFormulas::new(sink, Format::Ods))?
        } else {
            reader.read_sheets_into(sink)?
        };
        let defined_names = reader.read_defined_names()?;
        (
            sheets,
            defined_names,
            reader.read_name_conflicts()?,
            reader.read_hidden_sheets()?,
            reader.has_macros()?,
            reader.read_external_workbooks()?,
            reader.read_modified_date()?,
            reader.read_date_system()?,
            reader.read_settings()?,
            reader.read_properties()?,
        )
    } else {
        let mut reader = XlsxReader::new(&mut archive)?;
        (
            reader.read_sheets_into(sink)?,
            reader.read_defined_names()?,
            reader.read_name_conflicts()?,
            reader.read_hidden_sheets()?,
            reader.has_macros()?,
            reader.read_external_workbooks()?,
            reader.read_modified_date()?,
            reader.read_date_system()?,
            reader.read_settings()?,
            reader.read_properties()?,
        )
    };

    Ok(Workbook {
        path: Path::new("").to_path_buf(), // Default empty path for reader
        sheets,
        defined_names,
        name_conflicts,
        hidden_sheets,
        has_macros,
        external_workbooks,
        modified_date,
        date_system,
        settings,
        properties,
        file_size_bytes: 0, // Will be set by read_workbook when path is known
    })
}

/// The date system of the workbook in `reader`: the date of serial 0.
///
/// # Errors
///
/// [`Error::UnsupportedDateSystem`] when an ODS declares a null date that is not one of the three
/// date systems of `LibreOffice` Calc; otherwise when the file cannot be read or is not a workbook
/// of a supported format.
pub(crate) fn read_date_system_from_reader<R: std::io::Read + std::io::Seek>(
    reader: R,
) -> Result<DateSystem> {
    let mut archive = open_archive(reader)?;
    if detect::detect_archive(&mut archive)? == Format::Ods {
        OdsReader::new(&mut archive)?.read_date_system()
    } else {
        XlsxReader::new(&mut archive)?.read_date_system()
    }
}

/// The defined names of the workbook in `reader` and those it defines in more than one scope
/// (its format is read from its content). An XLSX is not read beyond its workbook part; an ODS has its names
/// after its tables, so it is read through.
///
/// # Errors
///
/// When the file cannot be read or is not a workbook of a supported format.
pub(crate) fn read_defined_names_from_reader<R: std::io::Read + std::io::Seek>(
    reader: R,
) -> Result<(HashMap<String, String>, Vec<NameConflict>)> {
    let mut archive = open_archive(reader)?;
    if detect::detect_archive(&mut archive)? == Format::Ods {
        let mut reader = OdsReader::new(&mut archive)?;
        Ok((reader.read_defined_names()?, reader.read_name_conflicts()?))
    } else {
        let mut reader = XlsxReader::new(&mut archive)?;
        Ok((reader.read_defined_names()?, reader.read_name_conflicts()?))
    }
}

/// Reads what `selection` asks for of the workbook in `reader`, handing the cells to `sink` (see
/// [`WorkbookReader::read_selection_into`]); its format is read from its content.
///
/// # Errors
///
/// When the file cannot be read or is not a workbook of a supported format.
pub(crate) fn read_selection_from_reader_into<R: std::io::Read + std::io::Seek>(
    reader: R,
    selection: &Selection,
    sink: &mut dyn CellSink,
) -> Result<Vec<SheetInfo>> {
    let mut archive = open_archive(reader)?;
    if detect::detect_archive(&mut archive)? == Format::Ods {
        OdsReader::new(&mut archive)?.read_selection_into(selection, sink)
    } else {
        XlsxReader::new(&mut archive)?.read_selection_into(selection, sink)
    }
}

/// The archive of a workbook: a zip. What is not a zip says so, and says it is not a workbook.
fn open_archive<R: std::io::Read + std::io::Seek>(reader: R) -> Result<ZipArchive<R>> {
    ZipArchive::new(reader).map_err(|error| match error {
        zip::result::ZipError::InvalidArchive(_) | zip::result::ZipError::UnsupportedArchive(_) => {
            Error::unsupported(format!(
                "not a zip archive, or a truncated one ({error}): not an XLSX, XLSM or ODS workbook"
            ))
        }
        other => Error::Archive(other),
    })
}

/// Tests of the readers on `tests/lint_book.{xlsx,ods}` (made by `tests/make_lint_book.py`), each file
/// judged on its own: what `LibreOffice` itself reports for that file is in
/// `tests/assets/oracle/lint_book.<ext>.oracle.json` (made by `make_oracle.py`), and an expectation is
/// that, or a decision of `SheetRS` said as such.
#[cfg(test)]
mod lint_book_oracle {
    use super::*;
    use std::io::Cursor;

    const FILES: [(&str, &[u8]); 2] = [
        (
            "xlsx",
            include_bytes!("../../../tests/lint_book.xlsx") as &[u8],
        ),
        (
            "ods",
            include_bytes!("../../../tests/lint_book.ods") as &[u8],
        ),
    ];

    fn read(bytes: &[u8], _extension: &str) -> Workbook {
        read_workbook_from_reader(Cursor::new(bytes)).unwrap()
    }

    /// What `LibreOffice` reports for the file: the object of its oracle.
    fn oracle(extension: &str) -> serde_json::Value {
        let path = format!(
            "{}/tests/assets/oracle/lint_book.{extension}.oracle.json",
            env!("CARGO_MANIFEST_DIR")
        );
        serde_json::from_slice(&std::fs::read(path).unwrap()).unwrap()
    }

    /// The oracle of one cell (0-based row and column).
    fn libre(oracle: &serde_json::Value, sheet: &str, row: u64, col: u64) -> serde_json::Value {
        oracle["cells"]
            .as_array()
            .unwrap()
            .iter()
            .find(|c| c["sheet"] == sheet && c["row"] == row && c["col"] == col)
            .unwrap_or_else(|| panic!("the oracle has no {sheet} {row},{col}"))
            .clone()
    }

    /// `LibreOffice` writes its keywords in capitals (`M/D/YYYY`) and an XLSX keeps the language of
    /// the format in front (`[$-409]mm/dd/yyyy`): the same format as `mm/dd/yyyy`.
    fn same_format(libre: &str, ours: &str) -> bool {
        let plain = |f: &str| f.replace("[$-409]", "").replace('\\', "").to_lowercase();
        plain(libre) == plain(ours)
    }

    /// A date is the number `LibreOffice` reports and has the format it reports (`Data!C3`, the
    /// serial 45001 with `MM/DD/YYYY`).
    #[test]
    fn a_date_is_the_number_and_the_format_libreoffice_reports() {
        for (extension, bytes) in FILES {
            let workbook = read(bytes, extension);
            let libre_cell = libre(&oracle(extension), "Data", 2, 2);
            let cell = workbook.sheet("Data").and_then(|s| s.cell(2, 2)).unwrap();
            let CellValue::Number(value) = cell.value else {
                panic!("{extension}: Data!C3 is not a number");
            };
            assert!(
                (value - libre_cell["value"].as_f64().unwrap()).abs() < 1e-9,
                "{extension}: {value}"
            );
            assert!(
                same_format(
                    libre_cell["format"].as_str().unwrap(),
                    cell.num_fmt.as_deref().unwrap()
                ),
                "{extension}: {:?}",
                cell.num_fmt
            );
        }
    }

    /// `Vulnerable!B1` holds `NOW()`: its value is the moment `LibreOffice` opened the file, so only
    /// its format (a date and time) is compared.
    #[test]
    fn the_format_of_a_date_and_time_is_the_one_libreoffice_reports() {
        for (extension, bytes) in FILES {
            let format = read(bytes, extension)
                .sheet("Vulnerable")
                .and_then(|s| s.cell(0, 1))
                .and_then(|c| c.num_fmt.clone())
                .unwrap();
            let libre_cell = libre(&oracle(extension), "Vulnerable", 0, 1);
            assert!(
                same_format(libre_cell["format"].as_str().unwrap(), &format),
                "{extension}: {format}"
            );
        }
    }

    /// The location of the external workbook is the one `LibreOffice` shows in the formula of
    /// `Links!B1` (`'file:///folder/prices.ods'#$Prices.B1`): the file names it relative to the
    /// document (`prices.ods` in the XLSX, `../prices.ods` in the ODS), and it is resolved against
    /// the folder of the document.
    #[test]
    fn the_external_workbook_is_resolved_against_the_document_as_libreoffice_does() {
        for (extension, _) in FILES {
            let path = format!(
                "{}/../tests/lint_book.{extension}",
                env!("CARGO_MANIFEST_DIR")
            );
            let folder =
                std::fs::canonicalize(format!("{}/../tests", env!("CARGO_MANIFEST_DIR"))).unwrap();
            let libre_formula = libre(&oracle(extension), "Links", 0, 1)["formula"]
                .as_str()
                .unwrap()
                .replace("<folder>", &folder.display().to_string());
            let libre_url = libre_formula.split('\'').nth(1).unwrap().to_string();
            let urls = crate::Book::from_file(path)
                .unwrap()
                .external_urls()
                .unwrap();
            assert_eq!(urls.len(), 1, "{extension}");
            assert_eq!(
                (urls[0].basename.as_str(), &urls[0].url),
                ("prices.ods", &libre_url)
            );
        }
    }

    /// The external workbook a file refers to is the one `LibreOffice` shows in the formula of
    /// `Links!B1` (`'file:///.../prices.ods'#$Prices.B1`), and the reader gives the formula with its
    /// number (`[1]Prices!B1`: the notation the analysis works on).
    #[test]
    fn the_external_workbook_is_the_one_libreoffice_shows() {
        for (extension, bytes) in FILES {
            let workbook = read(bytes, extension);
            let libre_formula = libre(&oracle(extension), "Links", 0, 1)["formula"]
                .as_str()
                .unwrap()
                .to_string();
            let url = libre_formula
                .split('\'')
                .nth(1)
                .expect("LibreOffice writes the file between quotes");
            assert_eq!(workbook.external_workbooks.len(), 1, "{extension}");
            assert_eq!(
                std::path::Path::new(&workbook.external_workbooks[0].path).file_name(),
                std::path::Path::new(url).file_name(),
                "{extension}"
            );
            let formula = workbook
                .sheet("Links")
                .and_then(|s| s.cell(0, 1))
                .and_then(Cell::as_formula)
                .unwrap();
            assert_eq!(formula, "[1]Prices!B1", "{extension}");
        }
    }

    /// Text with quotes, `<`, `>` and `&` (`&quot;`, `&lt;`, `&amp;` in the XML) reads as `LibreOffice`
    /// shows it: the formulas that use no reference to another sheet are the text it reports, and the
    /// ones that do are the same cell in Excel's notation (`$Inputs.B2` is `Inputs!B2`,
    /// `$'Sheet 2'.A1` is `'Sheet 2'!A1`).
    #[test]
    fn escaped_text_reads_as_libreoffice_shows_it() {
        for (extension, bytes) in FILES {
            let workbook = read(bytes, extension);
            let libre_book = oracle(extension);
            let cell = |row: u32| {
                workbook
                    .sheet("Escapes")
                    .and_then(|s| s.cell(row, 0))
                    .unwrap_or_else(|| panic!("{extension}: missing Escapes!A{}", row + 1))
            };
            let libre_formula = |row: u64| {
                libre(&libre_book, "Escapes", row, 0)["formula"]
                    .as_str()
                    .unwrap()
                    .to_string()
            };
            for row in [0, 2] {
                assert_eq!(
                    cell(row as u32).as_formula(),
                    Some(libre_formula(row).trim_start_matches('=')),
                    "{extension}: Escapes row {row}"
                );
            }
            assert_eq!(
                cell(1).as_formula(),
                Some(
                    libre_formula(1)
                        .trim_start_matches('=')
                        .replace("$Inputs.B2", "Inputs!B2")
                        .as_str()
                ),
                "{extension}"
            );
            assert_eq!(libre_formula(3), "=$'Sheet 2'.A1", "{extension}");
            assert_eq!(cell(3).as_formula(), Some("'Sheet 2'!A1"), "{extension}");
            let shown = libre(&libre_book, "Escapes", 0, 0)["string"]
                .as_str()
                .unwrap()
                .to_string();
            assert_eq!(shown, "Say \"hi\" <b>ok</b> & more");
            assert_eq!(
                cell(0).value,
                CellValue::Text(shown.as_str().into()),
                "{extension}"
            );
        }
    }
}
