//! The workbook: the one interface of the suite.
//!
//! A [`Book`] is what every tool takes as input and the only thing a caller holds: the linter reads
//! a model from it, the statistics stream it, a query reads the cells it asks for, the engine
//! calculates it and the editor streams it into another file. It records where the workbook
//! comes from (a file, or the bytes of one), the [`Format`] its content says it has and, with the
//! feature `write`, the operations still to be applied to it. Nothing is read until a tool reads it,
//! so a `Book` is cheap to make, to clone and to hand to several tools. It is not a navigable model
//! of the workbook: that is [`Workbook`], which only the tools that need every cell ask for
//! ([`Book::workbook`]).
//!
//! Each tool adds what it does to a `Book` in its own crate, as a trait to import
//! (`sheetrs_query::BookQuery`, `sheetrs_edit::BookEdit`, `sheetrs_compiler::BookCompile`,
//! `sheetrs_linter::BookLint`, `sheetrs_stats::BookStats`); the data they share is here.
//!
//! ```no_run
//! use sheetrs_io::Book;
//!
//! # fn main() -> sheetrs_io::Result<()> {
//! let book = Book::from_file("in.xlsx")?;
//! println!("{}: {} sheets", book.format(), book.workbook()?.sheets.len());
//! # Ok(()) }
//! ```

use crate::error::{Error, Result};
use crate::locale::{self, LocaleTag};
use crate::reader::{
    CellSink, Format, NameConflict, Selection, SheetInfo, Workbook, read_analysis_from_reader_into,
    read_date_system_from_reader, read_defined_name_definitions_from_reader,
    read_defined_names_from_reader, read_selection_from_reader_into, read_workbook_from_reader,
    read_workbook_from_reader_into,
};
use sheetrs_formula::DateSystem;
use std::collections::HashMap;
use std::fmt;
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::{Path, PathBuf};
use std::sync::Arc;

/// Where the workbook is.
#[derive(Debug, Clone, PartialEq, Eq)]
enum Origin {
    File(PathBuf),
    Bytes(Arc<[u8]>),
}

/// What the volatile functions give when the engine calculates the workbook: the serial number of
/// `NOW()` and `TODAY()` and the seed of `RAND()` and its relatives. A field that is `None` is the
/// system clock (the seed is taken from it). The names are those of `sheetc` (`--now`, `--seed`) and
/// of a `calculate` step.
///
/// It is a property of the [`Book`], like its locale ([`Book::with_volatile`]), and every tool that
/// builds or calculates takes it from there. The order, from the one that wins: the `now` and
/// `seed` of a single call (a `calculate` step, the options of a call), those of the book, the
/// configuration file (`now` and `seed` in `[global]` of `sheetrs.toml`,
/// [`Volatile::configured`]) and the clock. [`Volatile::or`] applies the order to two of them.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct Volatile {
    /// The serial number of `NOW()`; `TODAY()` is its integer part.
    pub now: Option<f64>,
    /// The seed of `RAND()`. Distinct seeds give distinct streams; `RAND` of `LibreOffice` cannot be
    /// seeded, so the seed is an extension of `SheetRS` with no `LibreOffice` counterpart.
    pub seed: Option<u64>,
}

impl Volatile {
    /// What `self` leaves open is taken from `base`: the fields of `self` win.
    ///
    /// ```
    /// use sheetrs_io::Volatile;
    ///
    /// let call = Volatile { now: Some(45000.0), seed: None };
    /// let book = Volatile { now: Some(1.0), seed: Some(7) };
    /// assert_eq!(call.or(book), Volatile { now: Some(45000.0), seed: Some(7) });
    /// ```
    #[must_use]
    pub fn or(self, base: Self) -> Self {
        Self {
            now: self.now.or(base.now),
            seed: self.seed.or(base.seed),
        }
    }

    /// The `now` and `seed` of the text of a configuration file (`sheetrs.toml`): the keys of
    /// `[global]`; absent, `None`.
    ///
    /// ```
    /// use sheetrs_io::Volatile;
    ///
    /// let text = "[global]\nnow = 45366.75\nseed = 7\n";
    /// let volatile = Volatile::from_config(text).unwrap();
    /// assert_eq!(volatile, Volatile { now: Some(45366.75), seed: Some(7) });
    /// ```
    ///
    /// # Errors
    ///
    /// [`Error::Config`] when the text is not TOML or a key has a value of the wrong type (`seed` is a
    /// non-negative integer, `now` a number).
    #[cfg(feature = "config")]
    pub fn from_config(text: &str) -> Result<Self> {
        #[derive(serde::Deserialize, Default)]
        struct Global {
            now: Option<f64>,
            seed: Option<u64>,
        }
        #[derive(serde::Deserialize, Default)]
        struct File {
            #[serde(default)]
            global: Global,
        }
        let file: File = toml::from_str(text).map_err(|e| Error::Config(e.to_string()))?;
        Ok(Self {
            now: file.global.now,
            seed: file.global.seed,
        })
    }

    /// The `now` and `seed` the configuration of a command line gives: the file `config` when the
    /// command names one (it must exist), else `sheetrs.toml` in the working directory when there is
    /// one; nothing when there is no file.
    ///
    /// # Errors
    ///
    /// [`Error::Open`] when the file cannot be read, and as [`from_config`](Self::from_config).
    #[cfg(feature = "config")]
    pub fn configured(config: Option<&Path>) -> Result<Self> {
        let default = Path::new(locale::CONFIG_FILE);
        let path = match config {
            Some(path) => path,
            None if default.is_file() => default,
            None => return Ok(Self::default()),
        };
        let text = std::fs::read_to_string(path).map_err(|source| Error::Open {
            path: path.to_path_buf(),
            source,
        })?;
        Self::from_config(&text).map_err(|error| Error::Context {
            message: path.display().to_string(),
            source: Box::new(error),
        })
    }
}

/// A workbook: where it comes from, its format, its locale and, with the feature `write`, the
/// operations pending on it.
#[derive(Debug, Clone, PartialEq)]
pub struct Book {
    origin: Origin,
    format: Format,
    locale: Option<LocaleTag>,
    volatile: Volatile,
    #[cfg(feature = "write")]
    operations: Vec<crate::write::Operation>,
    /// Whether a run calculates every formula (`calculateAll()` of `LibreOffice`) instead of keeping the
    /// results the file caches for the formulas nothing reaches.
    calculate_all: bool,
}

impl Book {
    /// A file. The format is read from its content (see [`Format::detect_reader`]), not from its
    /// extension: a mismatch between the two is not an error, the content wins. Nothing else is
    /// read.
    ///
    /// # Errors
    ///
    /// [`Error::Open`] when the file cannot be opened, and [`Error::Unsupported`] when it is not an
    /// XLSX, XLSM or ODS workbook.
    pub fn from_file(path: impl Into<PathBuf>) -> Result<Self> {
        let path = path.into();
        let file = File::open(&path).map_err(|source| Error::Open {
            path: path.clone(),
            source,
        })?;
        let format = Format::detect_reader(file).map_err(|error| match error {
            Error::Unsupported(what) => Error::unsupported(format!("'{}': {what}", path.display())),
            other => other,
        })?;
        Ok(Self::new(Origin::File(path), format))
    }

    /// The bytes of a workbook. The format is read from them (see [`Format::detect_reader`]).
    ///
    /// # Errors
    ///
    /// [`Error::Unsupported`] when they are not an XLSX, XLSM or ODS workbook.
    pub fn from_bytes(data: impl Into<Arc<[u8]>>) -> Result<Self> {
        let data = data.into();
        let format = Format::detect(&data)?;
        Ok(Self::new(Origin::Bytes(data), format))
    }

    /// The workbook `reader` gives, read to its end into memory (a stream, such as a pipe, cannot be
    /// read twice).
    ///
    /// # Errors
    ///
    /// When `reader` fails, and as [`from_bytes`](Self::from_bytes).
    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        Self::from_bytes(data)
    }

    fn new(origin: Origin, format: Format) -> Self {
        Self {
            origin,
            format,
            locale: None,
            volatile: Volatile::default(),
            #[cfg(feature = "write")]
            operations: Vec::new(),
            calculate_all: false,
        }
    }

    /// The same workbook with these values for the volatile functions (`NOW`, `TODAY`, `RAND` and its
    /// relatives) in every calculation that a tool makes of it: `calculate`, the recalculation that an
    /// edit leaves to do, the build of a program, `inspect` and `verify`. It is the way to initialise
    /// them, as [`with_locale`](Self::with_locale) is for the locale: `Book(..., now=, seed=)` in
    /// Python, `new Book(bytes, locale, {now, seed})` in WebAssembly, `--now` and `--seed` of
    /// `sheetrs edit`. The `now` and `seed` of a single call (a `calculate` step, the options of
    /// `inspect`) win over these, field by field.
    ///
    /// ```
    /// use sheetrs_io::{Book, Volatile};
    ///
    /// # fn main() -> sheetrs_io::Result<()> {
    /// # let bytes = include_bytes!("../../tests/lint_book.xlsx").to_vec();
    /// let book = Book::from_bytes(bytes)?.with_volatile(Volatile { now: Some(45366.75), seed: Some(7) });
    /// assert_eq!(book.volatile().seed, Some(7));
    /// # Ok(()) }
    /// ```
    #[must_use]
    pub fn with_volatile(mut self, volatile: Volatile) -> Self {
        self.volatile = volatile;
        self
    }

    /// The same workbook with the values of `other` for the volatile functions: for a book made from
    /// the result of an operation on `other`, which is the same workbook.
    #[must_use]
    pub fn with_volatile_of(mut self, other: &Self) -> Self {
        self.volatile = other.volatile;
        self
    }

    /// The same workbook with the calculation mode of `other` ([`with_calculate_all`](Self::with_calculate_all)):
    /// for a book made from `other`.
    #[must_use]
    pub fn with_calculate_all_of(mut self, other: &Self) -> Self {
        self.calculate_all = other.calculate_all;
        self
    }

    /// What the volatile functions give for this book ([`with_volatile`](Self::with_volatile)).
    #[must_use]
    pub fn volatile(&self) -> Volatile {
        self.volatile
    }

    /// The same workbook with this locale: the regional settings that give a text to a number and read
    /// a number from a text (`TEXT`, the conversion of a number to text, the `display` of a query, the
    /// dates a search reads), in every tool that takes the book. Without it the book has the locale its
    /// system gives (see [`default_locale`](Self::default_locale) and the [`locale`]
    /// module). A `locale` option of a single call still wins over it.
    ///
    /// The tag is not checked against the locales a tool has until a tool needs it.
    ///
    /// ```
    /// use sheetrs_io::Book;
    ///
    /// # fn main() -> sheetrs_io::Result<()> {
    /// # let bytes = include_bytes!("../../tests/lint_book.xlsx").to_vec();
    /// let book = Book::from_bytes(bytes)?.with_locale("es-CO".parse()?);
    /// assert_eq!(book.locale().map(ToString::to_string).as_deref(), Some("es-CO"));
    /// # Ok(()) }
    /// ```
    #[must_use]
    pub fn with_locale(mut self, locale: LocaleTag) -> Self {
        self.locale = Some(locale);
        self
    }

    /// The same workbook with the locale of `other`, if it was given one: for a book made from the
    /// result of an operation on `other`, which is the same workbook.
    #[must_use]
    pub fn with_locale_of(mut self, other: &Self) -> Self {
        self.locale.clone_from(&other.locale);
        self
    }

    /// The locale given to the book ([`with_locale`](Self::with_locale)); `None` when it was not given
    /// one. [`default_locale`](Self::default_locale) is the one it has then.
    #[must_use]
    pub fn locale(&self) -> Option<&LocaleTag> {
        self.locale.as_ref()
    }

    /// The locale the workbook has when none was given: the one of the system (see
    /// [`locale::system`]), for an XLSX and an ODS alike, as in `LibreOffice`, whose results change
    /// with the regional setting of the machine and not with the language an ODS stores. `None` is
    /// `en-US`: a system that has none, and WebAssembly. It reads no file.
    #[must_use]
    pub fn default_locale(&self) -> Option<LocaleTag> {
        locale::system()
    }

    /// The format of the workbook, as its content says it.
    #[must_use]
    pub fn format(&self) -> Format {
        self.format
    }

    /// The path of the file, `None` for bytes.
    #[must_use]
    pub fn path(&self) -> Option<&Path> {
        match &self.origin {
            Origin::File(path) => Some(path),
            Origin::Bytes(_) => None,
        }
    }

    /// The bytes the workbook was given as, `None` for a file.
    #[must_use]
    pub fn bytes(&self) -> Option<&Arc<[u8]>> {
        match &self.origin {
            Origin::File(_) => None,
            Origin::Bytes(data) => Some(data),
        }
    }

    /// Reads the whole workbook into the model the linter analyzes.
    ///
    /// # Errors
    ///
    /// When the file cannot be opened or the archive is not a valid workbook of its format.
    pub fn workbook(&self) -> Result<Workbook> {
        match &self.origin {
            Origin::File(path) => {
                let file = open(path)?;
                let size = std::fs::metadata(path).map_or(0, |m| m.len());
                let mut workbook = read_workbook_from_reader(file)?;
                workbook.path.clone_from(path);
                workbook.file_size_bytes = size;
                Ok(workbook)
            }
            Origin::Bytes(data) => read_workbook_from_reader(Cursor::new(data)),
        }
    }

    /// Streams the cells of the workbook into `sink` as they are read, without keeping them: the
    /// model it returns holds what is not a cell (sheet names, defined names, sizes).
    ///
    /// # Errors
    ///
    /// As [`workbook`](Self::workbook).
    pub fn workbook_into(&self, sink: &mut dyn CellSink) -> Result<Workbook> {
        match &self.origin {
            Origin::File(path) => {
                let file = open(path)?;
                let size = std::fs::metadata(path).map_or(0, |m| m.len());
                let mut workbook = read_workbook_from_reader_into(file, sink)?;
                workbook.path.clone_from(path);
                workbook.file_size_bytes = size;
                Ok(workbook)
            }
            Origin::Bytes(data) => read_workbook_from_reader_into(Cursor::new(data), sink),
        }
    }

    /// [`workbook_into`](Self::workbook_into) with the formulas in Excel's notation, the analysis
    /// form that the linter, the statistics and the compiler work on. An XLSX keeps its formulas so
    /// already; an ODS keeps `OpenFormula`, which this converts (see
    /// [`ExcelFormulas`](crate::reader::ExcelFormulas)).
    ///
    /// # Errors
    ///
    /// As [`workbook`](Self::workbook).
    pub fn analysis_into(&self, sink: &mut dyn CellSink) -> Result<Workbook> {
        match &self.origin {
            Origin::File(path) => {
                let file = open(path)?;
                let size = std::fs::metadata(path).map_or(0, |m| m.len());
                let mut workbook = read_analysis_from_reader_into(file, sink)?;
                workbook.path.clone_from(path);
                workbook.file_size_bytes = size;
                Ok(workbook)
            }
            Origin::Bytes(data) => read_analysis_from_reader_into(Cursor::new(data), sink),
        }
    }

    /// Reads what `selection` asks for and nothing else, handing the cells to `sink`: a sheet
    /// that is not selected is not parsed, the reading of a sheet stops after the last row a
    /// window reaches, and the properties not asked for are not built. Returns the sheets the
    /// reader went through (see
    /// [`WorkbookReader::read_selection_into`](crate::reader::WorkbookReader::read_selection_into)).
    ///
    /// # Errors
    ///
    /// When the file cannot be opened or the archive is not a valid workbook of its format.
    pub fn selection_into(
        &self,
        selection: &Selection,
        sink: &mut dyn CellSink,
    ) -> Result<Vec<SheetInfo>> {
        match &self.origin {
            Origin::File(path) => read_selection_from_reader_into(open(path)?, selection, sink),
            Origin::Bytes(data) => {
                read_selection_from_reader_into(Cursor::new(data), selection, sink)
            }
        }
    }

    /// The names of the sheets, in the order of the workbook (reads the start of every sheet only).
    ///
    /// # Errors
    ///
    /// As [`workbook`](Self::workbook).
    pub fn sheet_names(&self) -> Result<Vec<String>> {
        struct Nothing;
        impl CellSink for Nothing {
            fn cell(&mut self, _: u16, _: (u32, u32), _: crate::reader::Cell) {}
        }
        let selection = Selection::new(crate::reader::Properties::VALUE)
            .window_everywhere(crate::reader::Window::new(0, 0, 0, 0));
        let infos = self.selection_into(&selection, &mut Nothing)?;
        let mut names = vec![
            String::new();
            infos
                .iter()
                .map(|i| usize::from(i.index) + 1)
                .max()
                .unwrap_or(0)
        ];
        for info in infos {
            names[usize::from(info.index)] = info.name;
        }
        Ok(names)
    }

    /// The date system of the workbook: the date of serial 0 ([`DateSystem`], one of the three of
    /// `LibreOffice` Calc). An XLSX is [`Null1904`](DateSystem::Null1904) when it says `date1904`
    /// and [`Null1899`](DateSystem::Null1899) otherwise; an ODS says it in `table:null-date`.
    ///
    /// # Errors
    ///
    /// [`Error::UnsupportedDateSystem`] when an ODS declares another date, and as
    /// [`workbook`](Self::workbook).
    pub fn date_system(&self) -> Result<DateSystem> {
        match &self.origin {
            Origin::File(path) => read_date_system_from_reader(open(path)?),
            Origin::Bytes(data) => read_date_system_from_reader(Cursor::new(data)),
        }
    }

    /// The defined names of the workbook (name to definition, as [`Workbook::defined_names`]) and
    /// the ones it defines in more than one scope, without reading the cells of an XLSX.
    ///
    /// # Errors
    ///
    /// As [`workbook`](Self::workbook).
    pub fn defined_names(&self) -> Result<(HashMap<String, String>, Vec<NameConflict>)> {
        match &self.origin {
            Origin::File(path) => read_defined_names_from_reader(open(path)?),
            Origin::Bytes(data) => read_defined_names_from_reader(Cursor::new(data)),
        }
    }

    /// The definitions of the defined names (name to definition) and nothing else: [`defined_names`](Self::defined_names) also
    /// looks for the names that are defined in more than one scope, which an ODS makes read its content again.
    ///
    /// # Errors
    ///
    /// As [`workbook`](Self::workbook).
    pub fn defined_name_definitions(&self) -> Result<HashMap<String, String>> {
        match &self.origin {
            Origin::File(path) => read_defined_name_definitions_from_reader(open(path)?),
            Origin::Bytes(data) => read_defined_name_definitions_from_reader(Cursor::new(data)),
        }
    }

    /// The state of the rows of `sheets` as the workbook has them before any pending operation: for each sheet, the rows up to the last of
    /// `spans` (`(sheet, first row, last row)`, 0-based) that are hidden and the ones that hold a value or a formula, as bands
    /// `(first, last)` in order. An XLSX is read as a stream that stops after the last row asked for (no cell is built); an ODS is read
    /// as [`workbook`](Self::workbook) does. A sheet the workbook does not have gives no rows.
    ///
    /// # Errors
    ///
    /// As [`workbook`](Self::workbook).
    pub fn row_state(&self, spans: &[(&str, u32, u32)]) -> Result<HashMap<String, RowState>> {
        let mut last: HashMap<&str, u32> = HashMap::new();
        for &(sheet, _, to) in spans {
            let entry = last.entry(sheet).or_insert(to);
            *entry = (*entry).max(to);
        }
        if self.format == Format::Ods {
            let workbook = self.workbook()?;
            let mut out = HashMap::new();
            for sheet in workbook
                .sheets
                .iter()
                .filter(|s| last.contains_key(s.name.as_str()))
            {
                let mut rows: Vec<u32> = sheet
                    .cells
                    .iter()
                    .filter(|(_, cell)| crate::reader::is_used(cell))
                    .map(|((row, _), _)| row)
                    .collect();
                rows.sort_unstable();
                rows.dedup();
                let mut filled: Vec<(u32, u32)> = Vec::new();
                for row in rows {
                    match filled.last_mut() {
                        Some(band) if band.1 + 1 == row => band.1 = row,
                        _ => filled.push((row, row)),
                    }
                }
                out.insert(
                    sheet.name.clone(),
                    RowState {
                        hidden: sheet.hidden_rows.clone(),
                        filled,
                    },
                );
            }
            return Ok(out);
        }
        match &self.origin {
            Origin::File(path) => xlsx_row_state(open(path)?, &last),
            Origin::Bytes(data) => xlsx_row_state(Cursor::new(data), &last),
        }
    }

    /// The external workbooks the formulas refer to, in the order of the references that count
    /// them (`[1]` is the first, in an XLSX), with the URL `LibreOffice` shows for each:
    /// the address the formulas write it with.
    ///
    /// # Errors
    ///
    /// As [`workbook`](Self::workbook).
    pub fn external_urls(&self) -> Result<Vec<crate::reader::ExternalUrl>> {
        let workbook = self.workbook()?;
        Ok(workbook
            .external_workbooks
            .iter()
            .map(|w| crate::reader::ExternalUrl {
                basename: w.path.clone(),
                url: crate::reader::external_url(self.format, self.path(), &w.target),
                address: w.address.clone(),
            })
            .collect())
    }
}

#[cfg(feature = "write")]
impl Book {
    /// Adds an operation after the ones pending, which is how a crate adds what it does to a `Book`
    /// (the extension traits do it). An operation with nothing in it is not added. Each operation
    /// addresses the workbook **as the operations before it leave it**.
    pub fn push(&mut self, operation: crate::write::Operation) -> &mut Self {
        if !operation.is_empty() {
            self.operations.push(operation);
        }
        self
    }

    /// The pending operations, in the order they were added.
    #[must_use]
    pub fn operations(&self) -> &[crate::write::Operation] {
        &self.operations
    }

    /// Drops the pending operations.
    pub fn clear_operations(&mut self) -> &mut Self {
        self.operations.clear();
        self
    }

    /// The same workbook that, in every run that calculates (a calculation, what an edit leaves stale,
    /// a build), calculates every formula again (`true`), as `calculateAll()` of `LibreOffice` does, instead
    /// of what `LibreOffice` does by default (`false`): to keep the result the file caches for every
    /// formula that nothing reaches, and to calculate what an operation or an input reaches and what
    /// depends on it. A workbook whose results the file declares stale, or that reads cells at run time
    /// (`INDIRECT`, `OFFSET`), is calculated whole in both.
    #[must_use]
    pub fn with_calculate_all(mut self, calculate_all: bool) -> Self {
        self.calculate_all = calculate_all;
        self
    }

    /// Whether the run calculates every formula ([`with_calculate_all`](Self::with_calculate_all)).
    #[must_use]
    pub fn calculate_all(&self) -> bool {
        self.calculate_all
    }
}

impl Book {
    /// Whether the workbook asks the program that opens it to calculate every formula (`fullCalcOnLoad` in the
    /// `calcPr` of an XLSX, which the applications write when the results they leave are not current). An ODS has no
    /// such mark (`LibreOffice` writes none in its content, its settings or its metadata): it is `false`. This is
    /// the one place the mark is read: the compiler (`sheetc build` and `calculate`) takes a marked workbook as a run
    /// with `calculate_all`, and so does the editor, which is a calculation and a writer.
    ///
    /// # Errors
    ///
    /// When the file cannot be read or is not a zip archive.
    pub fn recalculation_marked(&self) -> Result<bool> {
        if self.format != Format::Xlsx {
            return Ok(false);
        }
        let mut xml = String::new();
        match (self.path(), self.bytes()) {
            (Some(path), _) => zip::ZipArchive::new(open(path)?)?
                .by_name("xl/workbook.xml")?
                .read_to_string(&mut xml)?,
            (None, Some(data)) => zip::ZipArchive::new(Cursor::new(data.as_ref()))?
                .by_name("xl/workbook.xml")?
                .read_to_string(&mut xml)?,
            (None, None) => unreachable!("a book is a file or bytes"),
        };
        Ok(workbook_marks_recalculation(&xml))
    }
}

/// Whether the `workbook.xml` of an XLSX has `fullCalcOnLoad` on in its `calcPr`.
fn workbook_marks_recalculation(xml: &str) -> bool {
    let mut reader = quick_xml::Reader::from_str(xml);
    while let Ok(event) = reader.read_event() {
        match event {
            quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e)
                if e.name().as_ref() == "calcPr" =>
            {
                return e.attributes().flatten().any(|a| {
                    a.key.as_ref() == "fullCalcOnLoad" && matches!(a.value.as_ref(), "1" | "true")
                });
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    false
}

/// The state of the rows of an XLSX (see [`Book::row_state`]).
fn xlsx_row_state<R: Read + std::io::Seek>(
    reader: R,
    last: &HashMap<&str, u32>,
) -> Result<HashMap<String, RowState>> {
    let mut archive = zip::ZipArchive::new(reader).map_err(Error::Archive)?;
    let mut out = HashMap::new();
    for (sheet, &to) in last {
        let (hidden, filled) =
            crate::reader::xlsx_parser::extract_row_state_from_xlsx(&mut archive, sheet, to)?;
        out.insert((*sheet).to_string(), RowState { hidden, filled });
    }
    Ok(out)
}

/// The rows of a sheet that are hidden and the ones that hold a value or a formula, as bands `(first, last)`, 0-based and
/// inclusive (see [`Book::row_state`]).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RowState {
    /// The hidden rows.
    pub hidden: Vec<(u32, u32)>,
    /// The rows with a value or a formula.
    pub filled: Vec<(u32, u32)>,
}

fn open(path: &Path) -> Result<File> {
    File::open(path).map_err(|source| Error::Open {
        path: path.to_path_buf(),
        source,
    })
}

impl TryFrom<&Path> for Book {
    type Error = Error;

    fn try_from(path: &Path) -> Result<Self> {
        Self::from_file(path)
    }
}

impl TryFrom<PathBuf> for Book {
    type Error = Error;

    fn try_from(path: PathBuf) -> Result<Self> {
        Self::from_file(path)
    }
}

impl TryFrom<&str> for Book {
    type Error = Error;

    fn try_from(path: &str) -> Result<Self> {
        Self::from_file(path)
    }
}

impl fmt::Display for Book {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.origin {
            Origin::File(path) => write!(f, "{}", path.display()),
            Origin::Bytes(data) => write!(f, "{} bytes of {}", data.len(), self.format),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::workbook_marks_recalculation;

    #[test]
    fn an_xlsx_that_asks_to_recalculate_when_it_is_opened_is_marked() {
        let marked = r#"<workbook><sheets/><calcPr calcId="1" fullCalcOnLoad="1"/></workbook>"#;
        let truthy = r#"<workbook><calcPr fullCalcOnLoad="true"></calcPr></workbook>"#;
        let plain = r#"<workbook><calcPr calcId="1"/></workbook>"#;
        assert!(workbook_marks_recalculation(marked));
        assert!(workbook_marks_recalculation(truthy));
        assert!(!workbook_marks_recalculation(plain));
        assert!(!workbook_marks_recalculation("<workbook/>"));
    }

    use super::*;
    use crate::reader::Format;
    use std::io::Write;
    use zip::write::SimpleFileOptions;

    const XLSX: &[u8] = include_bytes!("../../tests/lint_book.xlsx");
    const ODS: &[u8] = include_bytes!("../../tests/lint_book.ods");

    /// A zip with the given entries, for the shapes of file that are not workbooks.
    fn zip(entries: &[(&str, &str)]) -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        for (name, text) in entries {
            writer
                .start_file(*name, SimpleFileOptions::default())
                .unwrap();
            writer.write_all(text.as_bytes()).unwrap();
        }
        writer.finish().unwrap().into_inner()
    }

    fn content_types(workbook_type: &str) -> String {
        format!(
            r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{workbook_type}"/></Types>"#
        )
    }

    #[test]
    fn the_format_is_read_from_the_content_of_the_bytes_and_of_the_file() {
        assert_eq!(Format::detect(XLSX).unwrap(), Format::Xlsx);
        assert_eq!(Format::detect(ODS).unwrap(), Format::Ods);
        // A macro-enabled workbook says so in the content type of its main part.
        let macros = zip(&[(
            "[Content_Types].xml",
            &content_types("application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        )]);
        assert_eq!(Format::detect(&macros).unwrap(), Format::Xlsm);
        assert_eq!(Book::from_bytes(macros).unwrap().format(), Format::Xlsm);
        assert_eq!(Book::from_bytes(XLSX).unwrap().format(), Format::Xlsx);
        let tests = Path::new(env!("CARGO_MANIFEST_DIR")).join("../tests");
        assert_eq!(
            Book::from_file(tests.join("lint_book.ods"))
                .unwrap()
                .format(),
            Format::Ods
        );
        assert_eq!(Format::Ods.to_string(), "ods");
    }

    #[test]
    fn a_file_with_the_wrong_extension_is_what_its_content_says() {
        let folder = tempfile::tempdir().unwrap();
        for (name, bytes, format) in [
            ("an_ods.xlsx", ODS, Format::Ods),
            ("an_xlsx.ods", XLSX, Format::Xlsx),
            ("no_extension", XLSX, Format::Xlsx),
            ("odd.dat", ODS, Format::Ods),
        ] {
            let path = folder.path().join(name);
            std::fs::write(&path, bytes).unwrap();
            let book = Book::from_file(&path).unwrap();
            assert_eq!(book.format(), format, "{name}");
            // And it is read as what it is.
            assert!(!book.workbook().unwrap().sheets.is_empty(), "{name}");
        }
    }

    #[test]
    fn the_default_locale_is_the_one_of_the_system_in_both_formats() {
        // An ODS that states a language in its default cell style has the system's all the same:
        // `LibreOffice` does not read it.
        let ods = zip(&[
            ("mimetype", "application/vnd.oasis.opendocument.spreadsheet"),
            ("content.xml", "<office:document-content/>"),
            (
                "styles.xml",
                r#"<office:document-styles><office:styles><style:default-style style:family="table-cell"><style:text-properties fo:language="es" fo:country="CO"/></style:default-style></office:styles></office:document-styles>"#,
            ),
        ]);
        for bytes in [ods, XLSX.to_vec()] {
            assert_eq!(
                Book::from_bytes(bytes).unwrap().default_locale(),
                locale::system()
            );
        }
    }

    #[test]
    fn the_locale_given_to_a_book_is_kept_by_the_book_made_from_it() {
        let book = Book::from_bytes(XLSX).unwrap();
        assert_eq!(book.locale(), None);
        let book = book.with_locale("es-CO".parse().unwrap());
        assert_eq!(book.locale().unwrap().to_string(), "es-CO");
        assert_eq!(book.clone().locale(), book.locale());
        let next = Book::from_bytes(XLSX).unwrap().with_locale_of(&book);
        assert_eq!(next.locale(), book.locale());
        assert_ne!(next, Book::from_bytes(XLSX).unwrap());
    }

    #[test]
    fn the_volatile_given_to_a_book_is_kept_by_the_book_made_from_it() {
        let given = Volatile {
            now: Some(45366.75),
            seed: Some(7),
        };
        let book = Book::from_bytes(XLSX).unwrap();
        assert_eq!(book.volatile(), Volatile::default());
        let book = book.with_volatile(given);
        assert_eq!(book.clone().volatile(), given);
        let next = Book::from_bytes(XLSX).unwrap().with_volatile_of(&book);
        assert_eq!(next.volatile(), given);
    }

    #[test]
    fn what_a_call_leaves_open_is_the_books() {
        let call = Volatile {
            now: Some(1.0),
            seed: None,
        };
        let book = Volatile {
            now: Some(2.0),
            seed: Some(7),
        };
        assert_eq!(
            call.or(book),
            Volatile {
                now: Some(1.0),
                seed: Some(7)
            }
        );
        assert_eq!(
            Volatile::default().or(Volatile::default()),
            Volatile::default()
        );
    }

    #[cfg(feature = "config")]
    #[test]
    fn now_and_seed_are_read_from_global_of_a_configuration() {
        let read = |text: &str| Volatile::from_config(text);
        assert_eq!(
            read("[global]\nnow = 45366\nseed = 7\n").unwrap(),
            Volatile {
                now: Some(45366.0),
                seed: Some(7)
            }
        );
        assert_eq!(read("[global]\n").unwrap(), Volatile::default());
        assert_eq!(read("").unwrap(), Volatile::default());
        for bad in ["seed = -1", "seed = 1.5", "seed = \"7\"", "now = \"today\""] {
            assert!(
                matches!(
                    read(format!("[global]\n{bad}\n").as_str()),
                    Err(Error::Config(_))
                ),
                "{bad}"
            );
        }
    }

    #[cfg(feature = "config")]
    #[test]
    fn a_configuration_file_that_is_named_must_exist() {
        let dir = tempfile::tempdir().unwrap();
        let missing = dir.path().join("none.toml");
        assert!(matches!(
            Volatile::configured(Some(&missing)),
            Err(Error::Open { .. })
        ));
        let file = dir.path().join("sheetrs.toml");
        std::fs::write(&file, "[global]\nseed = 3\n").unwrap();
        assert_eq!(Volatile::configured(Some(&file)).unwrap().seed, Some(3));
    }

    #[test]
    fn a_book_reads_from_a_file_and_from_bytes_alike() {
        let book = Book::from_bytes(XLSX).unwrap();
        assert_eq!(book.path(), None);
        assert!(!book.workbook().unwrap().sheets.is_empty());
        assert!(book.to_string().contains("bytes of xlsx"));
        assert_eq!(book.clone(), book);
    }

    #[test]
    fn what_is_not_a_workbook_is_unsupported_and_says_what_it_is() {
        use crate::Error;
        let unsupported = |bytes: &[u8]| match Format::detect(bytes) {
            Err(Error::Unsupported(what)) => what,
            other => panic!("{other:?}"),
        };
        assert!(unsupported(b"").contains("empty input"));
        assert!(unsupported(b"not a zip at all").contains("not a zip archive"));
        assert!(unsupported(&XLSX[..XLSX.len() / 2]).contains("truncated"));
        assert!(unsupported(&zip(&[("readme.txt", "hello")])).contains("neither"));
        assert!(
            unsupported(&zip(&[
                ("mimetype", "application/vnd.oasis.opendocument.text"),
                ("content.xml", "<x/>"),
            ]))
            .contains("not an ODS spreadsheet")
        );
        let ct = |what: &str| content_types(what);
        assert!(
            unsupported(&zip(&[(
                "[Content_Types].xml",
                &ct(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
                )
            )]))
            .contains("XLTX")
        );
        assert!(
            unsupported(&zip(&[(
                "[Content_Types].xml",
                &ct("application/vnd.ms-excel.sheet.binary.macroEnabled.main+xml")
            )]))
            .contains("XLSB")
        );
        let word = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
        assert!(unsupported(&zip(&[("[Content_Types].xml", word)])).contains("Word document"));
    }

    #[test]
    fn the_errors_say_what_is_wrong_with_a_file() {
        use crate::Error;
        assert!(matches!(
            Book::from_file("does/not/exist.xlsx"),
            Err(Error::Open { path, .. }) if path.ends_with("exist.xlsx")
        ));
        // A missing file is an I/O failure that a caller can find by its book chain.
        let error = Book::from_file("does/not/exist.xlsx").unwrap_err();
        let io = std::error::Error::source(&error)
            .and_then(|e| e.downcast_ref::<std::io::Error>())
            .expect("an I/O error is the cause");
        assert_eq!(io.kind(), std::io::ErrorKind::NotFound);
        // A file that is not a workbook is unsupported, and the message names it.
        let folder = tempfile::tempdir().unwrap();
        let notes = folder.path().join("notes.txt");
        std::fs::write(&notes, "hello").unwrap();
        match Book::from_file(&notes) {
            Err(Error::Unsupported(what)) => assert!(what.contains("notes.txt"), "{what}"),
            other => panic!("{other:?}"),
        }
        // Bytes that are not a workbook are refused where they are given.
        assert!(matches!(
            Book::from_bytes(&b"not a zip"[..]),
            Err(Error::Unsupported(_))
        ));
    }
}
