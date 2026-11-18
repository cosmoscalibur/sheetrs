//! Streaming reader of XLSX workbooks into the sheetrs model.
//!
//! The workbook part is `xl/workbook.xml`, its relationships are in
//! `xl/_rels/workbook.xml.rels` and the shared strings in `xl/sharedStrings.xml`;
//! worksheets are read one at a time without loading the whole part into memory.
//!
//! Cell contents are read as `LibreOffice` loads them, including text that does not
//! fit the cell type. A formula whose cached result is missing is read with an empty value.

mod error;
mod lexical;
mod parts;
mod worksheet;
mod xml;

use std::io::{BufReader, Read, Seek};

use sheetrs_core::{Sheet, Workbook};
use zip::ZipArchive;
use zip::result::ZipError;

pub use error::Error;

use parts::{
    RELATIONSHIPS_PART, SHARED_STRINGS_PART, WORKBOOK_PART, parse_relationships,
    parse_shared_strings, parse_workbook,
};
use worksheet::{ValueContext, parse_worksheet};

const BUFFER_BYTES: usize = 64 * 1024;

/// Reads a workbook from an XLSX package.
///
/// # Errors
///
/// [`Error::Zip`] when the input is not a ZIP package, [`Error::MissingPart`]
/// when a part the workbook refers to is absent, and [`Error::Xml`] or
/// [`Error::Malformed`] when a part cannot be read.
pub fn read<R: Read + Seek>(source: R) -> Result<Workbook, Error> {
    let mut archive = ZipArchive::new(source)?;

    let workbook_part = parse_workbook(open(&mut archive, WORKBOOK_PART)?)?;
    let targets = parse_relationships(open(&mut archive, RELATIONSHIPS_PART)?)?;
    let shared_strings = match open_optional(&mut archive, SHARED_STRINGS_PART)? {
        Some(source) => parse_shared_strings(source)?,
        None => Vec::new(),
    };

    let context = ValueContext {
        shared_strings: &shared_strings,
        date_system: workbook_part.date_system,
    };
    let mut workbook = Workbook::new();
    workbook.set_date_system(workbook_part.date_system);
    for entry in workbook_part.sheets {
        let target = targets.get(&entry.relationship_id).ok_or_else(|| {
            Error::malformed(
                WORKBOOK_PART,
                format!("sheet {:?} refers to an unknown relationship", entry.name),
            )
        })?;
        let part = resolve_target(target);
        let mut sheet = Sheet::new(entry.name);
        sheet.set_visibility(entry.visibility);
        parse_worksheet(open(&mut archive, &part)?, &mut sheet, &context, &part)?;
        workbook.push_sheet(sheet);
    }
    Ok(workbook)
}

fn open<'a, R: Read + Seek>(
    archive: &'a mut ZipArchive<R>,
    name: &str,
) -> Result<BufReader<zip::read::ZipFile<'a, R>>, Error> {
    open_optional(archive, name)?.ok_or_else(|| Error::MissingPart(name.to_owned()))
}

fn open_optional<'a, R: Read + Seek>(
    archive: &'a mut ZipArchive<R>,
    name: &str,
) -> Result<Option<BufReader<zip::read::ZipFile<'a, R>>>, Error> {
    match archive.by_name(name) {
        Ok(file) => Ok(Some(BufReader::with_capacity(BUFFER_BYTES, file))),
        Err(ZipError::FileNotFound) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

/// The package path of a relationship target, which is relative to `xl/` unless it starts with `/`.
fn resolve_target(target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_owned();
    }
    let mut segments = vec!["xl"];
    for segment in target.split('/') {
        match segment {
            "." | "" => {}
            ".." => {
                segments.pop();
            }
            other => segments.push(other),
        }
    }
    segments.join("/")
}

#[cfg(test)]
mod tests {
    use super::resolve_target;

    #[test]
    fn targets_resolve_against_the_workbook_directory() {
        assert_eq!(
            resolve_target("worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_target("/xl/worksheets/sheet2.xml"),
            "xl/worksheets/sheet2.xml"
        );
        assert_eq!(
            resolve_target("../customXml/item1.xml"),
            "customXml/item1.xml"
        );
        assert_eq!(
            resolve_target("./worksheets//sheet3.xml"),
            "xl/worksheets/sheet3.xml"
        );
    }
}
