//! Streaming reader of XLSX workbooks into the sheetrs model.
//!
//! The workbook part is `xl/workbook.xml`, its relationships are in
//! `xl/_rels/workbook.xml.rels` and the shared strings in `xl/sharedStrings.xml`;
//! worksheets are read one at a time without loading the whole part into memory.
//!
//! Cell contents are read as `LibreOffice` loads them, including text that does not
//! fit the cell type. A formula whose cached result is missing is read with an empty value.

mod capture;
mod error;
mod lexical;
mod parallel;
mod parts;
mod vba;
mod worksheet;
mod write;
mod xml;

use std::io::{BufReader, Read, Seek};

use sheetrs_core::{OpaquePackage, OpaquePart, PackageFormat, SheetOpaque, Workbook};
use zip::ZipArchive;
use zip::result::ZipError;

pub use error::Error;
pub use write::write;

use parts::{
    RELATIONSHIPS_PART, SHARED_STRINGS_PART, WORKBOOK_PART, parse_external_link,
    parse_relationships, parse_shared_strings, parse_workbook,
};
use worksheet::{ValueContext, parse_worksheet};

const BUFFER_BYTES: usize = 64 * 1024;
/// The most bytes of a VBA project that the reader looks at.
const MAX_VBA_PROJECT_BYTES: u64 = 1 << 26;
/// Worksheets smaller than this are read on one thread even when more are allowed.
const PARALLEL_MIN_BYTES: u64 = 2 * 1024 * 1024;

/// Reads a workbook from an XLSX package.
///
/// # Errors
///
/// [`Error::Zip`] when the input is not a ZIP package, [`Error::MissingPart`]
/// when a part the workbook refers to is absent, and [`Error::Xml`] or
/// [`Error::Malformed`] when a part cannot be read.
pub fn read<R: Read + Seek>(source: R) -> Result<Workbook, Error> {
    read_with_workers(source, 1)
}

/// Reads a workbook like [`read`], parsing the large worksheets in blocks of rows on up to
/// `workers` threads. Prototype: it assumes every row has its `r` attribute.
///
/// # Errors
///
/// The same as [`read`].
pub fn read_parallel<R: Read + Seek>(source: R, workers: usize) -> Result<Workbook, Error> {
    read_with_workers(source, workers)
}

fn read_with_workers<R: Read + Seek>(source: R, workers: usize) -> Result<Workbook, Error> {
    let mut archive = ZipArchive::new(source)?;

    let workbook_part = parse_workbook(open(&mut archive, WORKBOOK_PART)?)?;
    let targets = parse_relationships(open(&mut archive, RELATIONSHIPS_PART)?, RELATIONSHIPS_PART)?;
    let shared_strings_part = open_optional(&mut archive, SHARED_STRINGS_PART)?;
    let has_shared_strings = shared_strings_part.is_some();
    let shared_strings = shared_strings_part
        .map(parse_shared_strings)
        .transpose()?
        .unwrap_or_default();

    let context = ValueContext {
        shared_strings: &shared_strings,
        date_system: workbook_part.date_system,
    };
    let mut workbook = Workbook::new();
    workbook.set_date_system(workbook_part.date_system);
    let mut sheet_parts = std::collections::HashSet::new();
    for entry in workbook_part.sheets {
        let target = targets
            .get(&entry.relationship_id)
            .map(|found| &found.target)
            .ok_or_else(|| {
                Error::malformed(
                    WORKBOOK_PART,
                    format!("sheet {:?} refers to an unknown relationship", entry.name),
                )
            })?;
        let part = resolve_target(target);
        let mut sheet = workbook.sheet_builder(entry.name.clone());
        sheet.set_visibility(entry.visibility);
        let file = open(&mut archive, &part)?;
        let frame = if workers > 1 && file.get_ref().size() >= PARALLEL_MIN_BYTES {
            parallel::parse_worksheet(file, &mut sheet, &context, &part, workers)?
        } else {
            Some(parse_worksheet(file, &mut sheet, &context, &part, true)?)
        };
        let frame = if let Some(frame) = frame {
            frame
        } else {
            // The blocks overlapped: only reading the part in order resolves the addresses.
            sheet.discard();
            sheet = workbook.sheet_builder(entry.name);
            sheet.set_visibility(entry.visibility);
            let file = open(&mut archive, &part)?;
            parse_worksheet(file, &mut sheet, &context, &part, true)?
        };
        sheet.set_opaque(SheetOpaque {
            frame,
            part: part.clone(),
            entry: entry.attributes,
        });
        sheet_parts.insert(part);
        sheet.finish_loaded();
    }
    for reference in &workbook_part.external_references {
        // A link that cannot be read names no document, and the workbook loads without it.
        let target = match external_target(&mut archive, &targets, reference) {
            Err(Error::Xml { .. } | Error::Malformed { .. }) => None,
            other => other?,
        };
        if let Some(target) = target {
            workbook.add_external_link(&target);
        }
    }
    list_vba_project(&mut archive, &targets, &mut workbook)?;
    workbook.set_opaque(OpaquePackage {
        format: PackageFormat::Xlsx,
        parts: unread_parts(&mut archive, &sheet_parts)?,
        frame: workbook_part.frame,
        date_system: workbook_part.date_system,
        shared_strings: has_shared_strings.then(|| SHARED_STRINGS_PART.to_owned()),
    });
    Ok(workbook)
}

/// The entries of the package that the model replaces when it writes: the workbook part, the
/// sheets and the shared strings. Every other entry is returned as it is.
fn unread_parts<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    sheet_parts: &std::collections::HashSet<String>,
) -> Result<Vec<OpaquePart>, Error> {
    let mut parts = Vec::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index)?;
        let name = file.name()?.into_owned();
        if file.is_dir()
            || name == WORKBOOK_PART
            || name == SHARED_STRINGS_PART
            || sheet_parts.contains(&name)
        {
            continue;
        }
        let mut bytes = Vec::with_capacity(usize::try_from(file.size()).unwrap_or(0));
        file.read_to_end(&mut bytes)
            .map_err(|error| Error::Zip(ZipError::Io(error)))?;
        parts.push(OpaquePart { name, bytes });
    }
    Ok(parts)
}

/// Adds to `workbook` the modules of its VBA project, if it has one.
fn list_vba_project<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    targets: &std::collections::HashMap<String, parts::Relationship>,
    workbook: &mut Workbook,
) -> Result<(), Error> {
    // Of several projects, the one with the first relationship id is read.
    let Some(relationship) = targets
        .iter()
        .filter(|(_, found)| found.kind.ends_with("/vbaProject"))
        .min_by_key(|(id, _)| id.as_str())
        .map(|(_, found)| found)
    else {
        return Ok(());
    };
    let part = resolve_target(&relationship.target);
    let Some(file) = open_optional(archive, &part)? else {
        return Ok(());
    };
    let mut bytes = Vec::new();
    file.take(MAX_VBA_PROJECT_BYTES + 1)
        .read_to_end(&mut bytes)
        .map_err(|error| Error::Zip(ZipError::Io(error)))?;
    // A project of that size is not read, as it is not one that people write.
    if u64::try_from(bytes.len()).is_ok_and(|length| length > MAX_VBA_PROJECT_BYTES) {
        return Ok(());
    }
    if let Some((library, modules)) = vba::project(&bytes) {
        for name in modules {
            workbook.add_macro_module(sheetrs_core::MacroModule {
                kind: sheetrs_core::MacroKind::Vba,
                library: library.clone(),
                name,
            });
        }
    }
    Ok(())
}

/// Where the external link that the workbook relationship `reference` names points, as its part
/// says.
fn external_target<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    targets: &std::collections::HashMap<String, parts::Relationship>,
    reference: &str,
) -> Result<Option<String>, Error> {
    let Some(target) = targets.get(reference) else {
        return Ok(None);
    };
    let part = resolve_target(&target.target);
    let Some(file) = open_optional(archive, &part)? else {
        return Ok(None);
    };
    let Some(book) = parse_external_link(file, &part)? else {
        return Ok(None);
    };
    let (directory, name) = part.rsplit_once('/').unwrap_or(("", &part));
    let relationships = format!("{directory}/_rels/{name}.rels");
    let Some(file) = open_optional(archive, &relationships)? else {
        return Ok(None);
    };
    // Only a relationship of type `externalLinkPath` names the document, and an empty target
    // names none.
    Ok(parse_relationships(file, &relationships)?
        .remove(&book)
        .filter(|found| found.kind.ends_with("/externalLinkPath") && !found.target.is_empty())
        .map(|found| found.target))
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
