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
    read_date_system_from_reader, read_defined_names_from_reader, read_selection_from_reader_into,
    read_workbook_from_reader, read_workbook_from_reader_into,
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

/// A workbook: where it comes from, its format, its locale and, with the feature `write`, the
/// operations pending on it.
#[derive(Debug, Clone, PartialEq)]
pub struct Book {
    origin: Origin,
    format: Format,
    locale: Option<LocaleTag>,
    #[cfg(feature = "write")]
    operations: Vec<crate::write::Operation>,
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
            #[cfg(feature = "write")]
            operations: Vec::new(),
        }
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
