//! Parsers of the workbook part, its relationships and the shared strings.

use std::collections::HashMap;
use std::io::BufRead;
use std::sync::Arc;

use quick_xml::Reader;
use quick_xml::events::Event;
use sheetrs_core::{DateSystem, Visibility};

use crate::Error;
use crate::lexical;
use crate::xml::{attribute, element_text, skip_element, unescape_xstring};

pub(crate) const WORKBOOK_PART: &str = "xl/workbook.xml";
pub(crate) const RELATIONSHIPS_PART: &str = "xl/_rels/workbook.xml.rels";
pub(crate) const SHARED_STRINGS_PART: &str = "xl/sharedStrings.xml";

/// A `<sheet>` entry of the workbook part.
pub(crate) struct SheetEntry {
    pub name: String,
    pub visibility: Visibility,
    pub relationship_id: String,
}

pub(crate) struct WorkbookPart {
    pub sheets: Vec<SheetEntry>,
    pub date_system: DateSystem,
}

pub(crate) fn parse_workbook<R: BufRead>(source: R) -> Result<WorkbookPart, Error> {
    let part = WORKBOOK_PART;
    let mut reader = Reader::from_reader(source);
    let mut buf = Vec::new();
    let mut workbook = WorkbookPart {
        sheets: Vec::new(),
        date_system: DateSystem::Date1900,
    };
    let mut complete = false;
    loop {
        buf.clear();
        match reader
            .read_event_into(&mut buf)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Empty(element) if element.local_name().as_ref() == "workbook" => {
                complete = true;
            }
            Event::Start(element) | Event::Empty(element) => match element.local_name().as_ref() {
                "workbookPr" => {
                    // The last properties element wins.
                    let on = attribute(&element, "date1904", part)?
                        .is_some_and(|text| lexical::date_1904_flag(&text));
                    workbook.date_system = if on {
                        DateSystem::Date1904
                    } else {
                        DateSystem::Date1900
                    };
                }
                "sheet" => workbook.sheets.push(sheet_entry(&element, part)?),
                _ => {}
            },
            Event::End(element) if element.local_name().as_ref() == "workbook" => complete = true,
            Event::Eof if complete => return Ok(workbook),
            Event::Eof => return Err(Error::malformed(part, "unexpected end of file")),
            _ => {}
        }
    }
}

fn sheet_entry(
    element: &quick_xml::events::BytesStart<'_>,
    part: &str,
) -> Result<SheetEntry, Error> {
    let required = |name: &str| {
        attribute(element, name, part)?
            .ok_or_else(|| Error::malformed(part, format!("a sheet has no {name} attribute")))
    };
    let visibility = match attribute(element, "state", part)?.as_deref() {
        None | Some("visible") => Visibility::Visible,
        Some("hidden") => Visibility::Hidden,
        Some("veryHidden") => Visibility::VeryHidden,
        Some(other) => {
            return Err(Error::malformed(
                part,
                format!("unknown sheet state {other:?}"),
            ));
        }
    };
    Ok(SheetEntry {
        name: required("name")?,
        visibility,
        relationship_id: required("id")?,
    })
}

/// Maps each relationship id to its target.
pub(crate) fn parse_relationships<R: BufRead>(source: R) -> Result<HashMap<String, String>, Error> {
    let part = RELATIONSHIPS_PART;
    let mut reader = Reader::from_reader(source);
    let mut buf = Vec::new();
    let mut targets = HashMap::new();
    let mut complete = false;
    loop {
        buf.clear();
        match reader
            .read_event_into(&mut buf)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == "Relationship" =>
            {
                let (Some(id), Some(target)) = (
                    attribute(&element, "Id", part)?,
                    attribute(&element, "Target", part)?,
                ) else {
                    return Err(Error::malformed(part, "a relationship has no Id or Target"));
                };
                targets.insert(id, target);
            }
            Event::Empty(element) if element.local_name().as_ref() == "Relationships" => {
                complete = true;
            }
            Event::End(element) if element.local_name().as_ref() == "Relationships" => {
                complete = true;
            }
            Event::Eof if complete => return Ok(targets),
            Event::Eof => return Err(Error::malformed(part, "unexpected end of file")),
            _ => {}
        }
    }
}

/// The shared strings in order, with rich-text runs joined and phonetic runs dropped.
pub(crate) fn parse_shared_strings<R: BufRead>(source: R) -> Result<Vec<Arc<str>>, Error> {
    let part = SHARED_STRINGS_PART;
    let mut reader = Reader::from_reader(source);
    let mut buf = Vec::new();
    let mut strings = Vec::new();
    let mut current: Option<String> = None;
    let mut complete = false;
    loop {
        buf.clear();
        match reader
            .read_event_into(&mut buf)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Start(element) => match (element.local_name().as_ref(), current.as_mut()) {
                ("si", _) => current = Some(String::new()),
                ("rPh", Some(_)) => skip_element(&mut reader, &element, part)?,
                ("t", Some(text)) => {
                    let mut scratch = Vec::new();
                    text.push_str(&unescape_xstring(&element_text(
                        &mut reader,
                        &mut scratch,
                        part,
                    )?));
                }
                _ => {}
            },
            Event::Empty(element) if element.local_name().as_ref() == "si" => {
                strings.push(Arc::from(""));
            }
            Event::End(element) if element.local_name().as_ref() == "si" => {
                let text = current.take().unwrap_or_default();
                strings.push(Arc::from(lexical::normalize_line_breaks(&text).as_ref()));
            }
            Event::Empty(element) if element.local_name().as_ref() == "sst" => complete = true,
            Event::End(element) if element.local_name().as_ref() == "sst" => complete = true,
            Event::Eof if complete => return Ok(strings),
            Event::Eof => return Err(Error::malformed(part, "unexpected end of file")),
            _ => {}
        }
    }
}
