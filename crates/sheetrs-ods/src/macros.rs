//! The Basic libraries of a package, listed without being read.

use std::collections::HashMap;
use std::collections::hash_map::Entry;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use sheetrs_core::{MacroKind, MacroModule, Workbook};
use zip::ZipArchive;
use zip::result::ZipError;

use crate::Error;

const LIBRARIES_PART: &str = "Basic/script-lc.xml";

/// Adds to `workbook` the modules of the libraries that `Basic/script-lc.xml` lists, as each
/// library's `script-lb.xml` lists them.
///
/// `LibreOffice` holds no library at all when a name is listed twice, and cannot load a package
/// whose names differ only in case. A library with a module whose file is missing holds no
/// modules.
pub(crate) fn list<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    workbook: &mut Workbook,
) -> Result<(), Error> {
    // A part that is not well-formed lists nothing, and the package loads without it.
    let Some(libraries) = readable(names(archive, LIBRARIES_PART, "library"))? else {
        return Ok(());
    };
    if repeated(&libraries, LIBRARIES_PART)? {
        return Ok(());
    }
    let mut found = Vec::new();
    for library in libraries {
        let part = format!("Basic/{library}/script-lb.xml");
        let Some(modules) = readable(names(archive, &part, "element"))? else {
            continue;
        };
        if repeated(&modules, &part)? {
            return Ok(());
        }
        let complete = modules.iter().all(|module| {
            archive
                .index_for_name(&format!("Basic/{library}/{module}.xml"))
                .is_some()
        });
        if complete {
            for module in modules {
                found.push(MacroModule {
                    kind: MacroKind::Basic,
                    library: library.clone(),
                    name: module,
                });
            }
        }
    }
    for module in found {
        workbook.add_macro_module(module);
    }
    Ok(())
}

/// `None` when the part is not well-formed XML, which `LibreOffice` reads as a part without
/// entries.
fn readable(result: Result<Vec<String>, Error>) -> Result<Option<Vec<String>>, Error> {
    match result {
        Err(Error::Xml { .. } | Error::Malformed { .. }) => Ok(None),
        other => other.map(Some),
    }
}

/// Whether a name is in `names` twice. Two names that differ only in case make the package
/// unreadable, whatever else is repeated.
fn repeated(names: &[String], part: &str) -> Result<bool, Error> {
    let mut by_lowercase: HashMap<String, &String> = HashMap::new();
    let mut repeats = false;
    for name in names {
        match by_lowercase.entry(name.to_lowercase()) {
            Entry::Vacant(slot) => {
                slot.insert(name);
            }
            Entry::Occupied(previous) if *previous.get() == name => repeats = true,
            Entry::Occupied(previous) => {
                return Err(Error::malformed(
                    part,
                    format!("{:?} and {name:?} differ only in case", previous.get()),
                ));
            }
        }
    }
    Ok(repeats)
}

/// The `library:name` of each `element` of the part, which is empty when the package has none.
fn names<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    part: &str,
    element: &str,
) -> Result<Vec<String>, Error> {
    let file = match archive.by_name(part) {
        Ok(file) => file,
        Err(ZipError::FileNotFound) => return Ok(Vec::new()),
        Err(error) => return Err(error.into()),
    };
    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buffer = Vec::new();
    let mut found = Vec::new();
    loop {
        buffer.clear();
        match reader
            .read_event_into(&mut buffer)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Start(tag) | Event::Empty(tag) if tag.local_name().as_ref() == element => {
                let mut attributes = tag.attributes();
                attributes.with_checks(false);
                for attribute in attributes {
                    let attribute =
                        attribute.map_err(|error| Error::malformed(part, error.to_string()))?;
                    if attribute.key.local_name().as_ref() == "name" {
                        found.push(
                            attribute
                                .normalized_value(XmlVersion::Implicit1_0)
                                .map_err(|source| Error::xml(part, source))?
                                .into_owned(),
                        );
                    }
                }
            }
            Event::Eof => return Ok(found),
            _ => {}
        }
    }
}
