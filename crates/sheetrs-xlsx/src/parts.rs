//! Parsers of the workbook part, its relationships and the shared strings.

use std::collections::HashMap;
use std::io::BufRead;
use std::sync::Arc;

use quick_xml::Reader;
use quick_xml::events::Event;
use sheetrs_core::{DateSystem, Frame, Visibility};

use crate::Error;
use crate::capture::{Tee, without_start_tag};
use crate::lexical;
use crate::xml::{attribute, element_text, raw_attributes, skip_element, unescape_xstring};

pub(crate) const WORKBOOK_PART: &str = "xl/workbook.xml";
pub(crate) const RELATIONSHIPS_PART: &str = "xl/_rels/workbook.xml.rels";
pub(crate) const SHARED_STRINGS_PART: &str = "xl/sharedStrings.xml";

/// A relationship of a part: what it is, and what it points at.
pub(crate) struct Relationship {
    pub kind: String,
    pub target: String,
}

/// A `<sheet>` entry of the workbook part.
pub(crate) struct SheetEntry {
    pub name: String,
    pub visibility: Visibility,
    pub relationship_id: String,
    /// The attributes of the entry other than its name and state, as written.
    pub attributes: String,
}

/// A `<definedName>` element as written.
pub(crate) struct DefinedName {
    pub name: Option<String>,
    pub local_sheet: Option<String>,
    pub hidden: Option<String>,
    /// The text of the element, which is the expression.
    pub text: String,
    /// The attributes as written, to write the element back unchanged.
    pub attributes: String,
}

pub(crate) struct WorkbookPart {
    pub defined_names: Vec<DefinedName>,
    pub sheets: Vec<SheetEntry>,
    pub date_system: DateSystem,
    /// The relationship ids of the external links, in order.
    pub external_references: Vec<String>,
    /// The part around the list of sheets.
    pub frame: Frame,
}

/// The index of the sheet the workbook opens on, as `LibreOffice` reads the `activeTab`
/// attribute: the integer it starts with when that is a sheet, and the first sheet otherwise.
fn active_sheet(attribute: Option<&str>, sheets: usize) -> usize {
    attribute
        .map(sheetrs_io::integer_prefix)
        .and_then(|index| usize::try_from(index).ok())
        .filter(|index| *index < sheets)
        .unwrap_or(0)
}

pub(crate) fn parse_workbook<R: BufRead>(source: R) -> Result<WorkbookPart, Error> {
    let part = WORKBOOK_PART;
    let mut reader = Reader::from_reader(Tee::new(source, true));
    let mut buf = Vec::new();
    let mut head = Vec::new();
    let mut prefixed = false;
    let mut scratch = Vec::new();
    let mut workbook = WorkbookPart {
        defined_names: Vec::new(),
        sheets: Vec::new(),
        date_system: DateSystem::Date1900,
        external_references: Vec::new(),
        frame: Frame::default(),
    };
    let mut complete = false;
    let mut active_tab: Option<String> = None;
    loop {
        buf.clear();
        match reader
            .read_event_into(&mut buf)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Empty(element) if element.local_name().as_ref() == "workbook" => {
                complete = true;
            }
            Event::Start(element) if element.local_name().as_ref() == "sheets" => {
                prefixed = element.name().as_ref().contains(':');
                head = without_start_tag(reader.get_mut().pause(), element.name().as_ref());
            }
            Event::Empty(element) if element.local_name().as_ref() == "sheets" => {
                head = without_start_tag(reader.get_mut().pause(), element.name().as_ref());
                reader.get_mut().resume();
            }
            Event::End(element) if element.local_name().as_ref() == "sheets" => {
                reader.get_mut().resume();
            }
            Event::Start(element) if element.local_name().as_ref() == "definedName" => {
                let text = element_text(&mut reader, &mut scratch, part)?;
                workbook
                    .defined_names
                    .push(defined_name(&element, text, part)?);
            }
            Event::Empty(element) if element.local_name().as_ref() == "definedName" => {
                workbook
                    .defined_names
                    .push(defined_name(&element, String::new(), part)?);
            }
            Event::Start(element) | Event::Empty(element) => match element.local_name().as_ref() {
                "workbookPr" => {
                    // The last properties element wins.
                    let on = attribute(&element, "date1904", part)?
                        .is_some_and(|text| lexical::workbook_flag(&text));
                    workbook.date_system = if on {
                        DateSystem::Date1904
                    } else {
                        DateSystem::Date1900
                    };
                }
                "workbookView" if active_tab.is_none() => {
                    active_tab = Some(attribute(&element, "activeTab", part)?.unwrap_or_default());
                }
                "sheet" => workbook.sheets.push(sheet_entry(&element, part)?),
                "externalReference" => {
                    if let Some(id) = attribute(&element, "id", part)? {
                        workbook.external_references.push(id);
                    }
                }
                _ => {}
            },
            Event::End(element) if element.local_name().as_ref() == "workbook" => complete = true,
            Event::Eof if complete => {
                // The sheet the workbook opens on is shown, whatever its state says.
                let active = active_sheet(active_tab.as_deref(), workbook.sheets.len());
                if let Some(entry) = workbook.sheets.get_mut(active) {
                    entry.visibility = Visibility::Visible;
                }
                workbook.frame = Frame {
                    head,
                    tail: reader.get_mut().take(),
                    prefixed,
                };
                return Ok(workbook);
            }
            Event::Eof => return Err(Error::malformed(part, "unexpected end of file")),
            _ => {}
        }
    }
}

fn defined_name(
    element: &quick_xml::events::BytesStart<'_>,
    text: String,
    part: &str,
) -> Result<DefinedName, Error> {
    Ok(DefinedName {
        name: attribute(element, "name", part)?,
        local_sheet: attribute(element, "localSheetId", part)?,
        hidden: attribute(element, "hidden", part)?,
        text,
        attributes: raw_attributes(element, &[], part)?,
    })
}

fn sheet_entry(
    element: &quick_xml::events::BytesStart<'_>,
    part: &str,
) -> Result<SheetEntry, Error> {
    let required = |name: &str| {
        attribute(element, name, part)?
            .ok_or_else(|| Error::malformed(part, format!("a sheet has no {name} attribute")))
    };
    // Any other state, in any case, is read as visible.
    let visibility = match attribute(element, "state", part)?.as_deref() {
        Some("hidden") => Visibility::Hidden,
        Some("veryHidden") => Visibility::VeryHidden,
        _ => Visibility::Visible,
    };
    Ok(SheetEntry {
        name: required("name")?,
        visibility,
        relationship_id: required("id")?,
        attributes: raw_attributes(element, &["name", "state"], part)?,
    })
}

/// Maps each relationship id of the part `part` to its target.
pub(crate) fn parse_relationships<R: BufRead>(
    source: R,
    part: &str,
) -> Result<HashMap<String, Relationship>, Error> {
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
                let kind = attribute(&element, "Type", part)?.unwrap_or_default();
                targets.insert(id, Relationship { kind, target });
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
/// The shared strings of a workbook.
pub(crate) struct SharedStrings {
    /// The text of each string, without formatting.
    pub plain: Vec<Arc<str>>,
    /// The strings that have formatting runs, by index, as the part writes them.
    pub formatted: HashMap<usize, String>,
}

pub(crate) fn parse_shared_strings<R: BufRead>(source: R) -> Result<SharedStrings, Error> {
    let part = SHARED_STRINGS_PART;
    let mut reader = Reader::from_reader(Tee::new(source, true));
    let mut buf = Vec::new();
    let mut strings = SharedStrings {
        plain: Vec::new(),
        formatted: HashMap::new(),
    };
    let mut current: Option<String> = None;
    let mut start_tag = String::new();
    let mut has_runs = false;
    let mut prefixed = false;
    let mut complete = false;
    loop {
        buf.clear();
        match reader
            .read_event_into(&mut buf)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Start(element) => match (element.local_name().as_ref(), current.as_mut()) {
                ("si", _) => {
                    prefixed |= element.name().as_ref().contains(':');
                    current = Some(String::new());
                    has_runs = false;
                    // What is kept of the element is its start tag and what follows it.
                    reader.get_mut().take();
                    let attributes = raw_attributes(&element, &[], part)?;
                    start_tag = if attributes.is_empty() {
                        "<si>".to_owned()
                    } else {
                        format!("<si {attributes}>")
                    };
                }
                ("r", Some(_)) => has_runs = true,
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
                strings.plain.push(Arc::from(""));
            }
            Event::End(element) if element.local_name().as_ref() == "si" => {
                let text = current.take().unwrap_or_default();
                if has_runs {
                    let inside = String::from_utf8_lossy(&reader.get_mut().take()).into_owned();
                    strings
                        .formatted
                        .insert(strings.plain.len(), format!("{start_tag}{inside}"));
                }
                strings
                    .plain
                    .push(Arc::from(lexical::normalize_line_breaks(&text).as_ref()));
            }
            Event::Empty(element) if element.local_name().as_ref() == "sst" => complete = true,
            Event::End(element) if element.local_name().as_ref() == "sst" => complete = true,
            Event::Eof if complete => {
                // Formatted text is written under unprefixed elements, which a prefixed part
                // does not have.
                if prefixed {
                    strings.formatted.clear();
                }
                return Ok(strings);
            }
            Event::Eof => return Err(Error::malformed(part, "unexpected end of file")),
            _ => {}
        }
    }
}

/// The relationship id of the document an external link part refers to, if it names one. What
/// comes after the `externalBook` element is not read, so a part that goes wrong later still
/// names its document, as it does for `LibreOffice`.
pub(crate) fn parse_external_link<R: BufRead>(
    source: R,
    part: &str,
) -> Result<Option<String>, Error> {
    let mut reader = Reader::from_reader(source);
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader
            .read_event_into(&mut buf)
            .map_err(|source| Error::xml(part, source))?
        {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == "externalBook" =>
            {
                return attribute(&element, "id", part);
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
    }
}
