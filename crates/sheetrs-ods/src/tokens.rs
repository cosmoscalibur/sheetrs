//! A stream of the elements of `content.xml`, with names resolved to their namespaces.

use std::io::BufRead;

use quick_xml::NsReader;
use quick_xml::XmlVersion;
use quick_xml::escape::resolve_xml_entity;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::ResolveResult;

use crate::Error;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const CALCEXT: &str = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";

/// The elements the reader tells apart; every other element is [`Tag::Other`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Tag {
    Table,
    Row,
    /// A group of rows that adds nothing to the cells: `table-row-group`, `table-header-rows`
    /// and `table-rows`.
    RowContainer,
    Cell,
    CoveredCell,
    Paragraph,
    Space,
    Tab,
    LineBreak,
    /// An element whose text is not part of the paragraph around it.
    Aside,
    Span,
    Link,
    Style,
    TableProperties,
    NullDate,
    /// The list of the names of a document or of a table.
    NamedExpressions,
    /// A name that stands for a range.
    NamedRange,
    /// A name that stands for an expression.
    NamedExpression,
    Other,
}

/// The values of the attributes the reader uses, by the element they were read from.
#[derive(Debug, Default)]
pub(crate) struct Attrs {
    pub value_type: Option<String>,
    pub value: Option<String>,
    pub string_value: Option<String>,
    pub boolean_value: Option<String>,
    pub date_value: Option<String>,
    pub time_value: Option<String>,
    pub calcext_type: Option<String>,
    pub formula: Option<String>,
    pub columns_repeated: Option<String>,
    pub rows_repeated: Option<String>,
    pub matrix_columns: Option<String>,
    pub matrix_rows: Option<String>,
    pub name: Option<String>,
    pub style_name: Option<String>,
    pub display: Option<String>,
    pub family: Option<String>,
    pub count: Option<String>,
    /// The `table:date-value` of a `table:null-date`.
    pub null_date: Option<String>,
    /// The `table:cell-range-address` of a named range.
    pub range_address: Option<String>,
    /// The `table:expression` of a named expression.
    pub expression: Option<String>,
}

pub(crate) enum Token {
    Start(Tag, Attrs),
    Empty(Tag, Attrs),
    End(Tag),
    Text(String),
    Eof,
}

pub(crate) struct Tokens<R: BufRead> {
    reader: NsReader<R>,
    buf: Vec<u8>,
    part: String,
    /// The elements that are open.
    depth: usize,
    /// Whether the root element has ended, after which nothing but whitespace may follow.
    root_closed: bool,
}

impl<R: BufRead> Tokens<R> {
    pub(crate) fn new(source: R, part: &str) -> Self {
        let reader = NsReader::from_reader(source);
        Self {
            reader,
            buf: Vec::new(),
            part: part.to_owned(),
            depth: 0,
            root_closed: false,
        }
    }

    pub(crate) fn part(&self) -> &str {
        &self.part
    }

    pub(crate) fn next(&mut self) -> Result<Token, Error> {
        self.buf.clear();
        let (resolved, event) = self
            .reader
            .read_resolved_event_into(&mut self.buf)
            .map_err(|source| Error::xml(&self.part, source))?;
        let tag = match &event {
            Event::Start(element) | Event::Empty(element) => {
                classify(&resolved, element.local_name().as_ref())
            }
            Event::End(element) => classify(&resolved, element.local_name().as_ref()),
            _ => Tag::Other,
        };
        // `LibreOffice` rejects a document that is cut short, one with end tags that do not match
        // and one with anything but whitespace after the root element.
        let after_root = || Error::malformed(&self.part, "content after the root element");
        Ok(match event {
            Event::Start(element) => {
                if self.root_closed {
                    return Err(after_root());
                }
                self.depth += 1;
                Token::Start(tag, attributes(&self.reader, &self.part, &element)?)
            }
            Event::Empty(element) => {
                if self.root_closed {
                    return Err(after_root());
                }
                self.root_closed = self.depth == 0;
                Token::Empty(tag, attributes(&self.reader, &self.part, &element)?)
            }
            Event::End(_) => {
                self.depth = self.depth.saturating_sub(1);
                self.root_closed = self.depth == 0;
                Token::End(tag)
            }
            Event::Text(content) => {
                let text = content.xml10_content().into_owned();
                if self.root_closed && !text.trim().is_empty() {
                    return Err(after_root());
                }
                Token::Text(text)
            }
            Event::CData(content) => Token::Text(content.xml10_content().into_owned()),
            Event::GeneralRef(reference) => Token::Text(reference_text(&reference, &self.part)?),
            Event::Eof if self.depth > 0 => {
                return Err(Error::malformed(&self.part, "unexpected end of file"));
            }
            Event::Eof => Token::Eof,
            _ => Token::Text(String::new()),
        })
    }

    /// Skips what is left of the element whose start was just read.
    pub(crate) fn skip(&mut self) -> Result<(), Error> {
        let mut depth = 0_usize;
        loop {
            match self.next()? {
                Token::Start(..) => depth += 1,
                Token::End(_) => {
                    if depth == 0 {
                        return Ok(());
                    }
                    depth -= 1;
                }
                Token::Eof => return Ok(()),
                _ => {}
            }
        }
    }
}

/// The values of the attributes of `element` that the reader uses.
fn attributes<R: BufRead>(
    reader: &NsReader<R>,
    part: &str,
    element: &BytesStart<'_>,
) -> Result<Attrs, Error> {
    let mut attrs = Attrs::default();
    let mut iterator = element.attributes();
    // A repeated attribute is not an error: `LibreOffice` reads it too.
    iterator.with_checks(false);
    for attribute in iterator {
        let attribute = attribute.map_err(|error| Error::malformed(part, error.to_string()))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let ResolveResult::Bound(namespace) = resolved else {
            continue;
        };
        let value = attribute
            .normalized_value(XmlVersion::Implicit1_0)
            .map_err(|source| Error::xml(part, source))?
            .into_owned();
        let slot = match (namespace.0, local.as_ref()) {
            (OFFICE, "value-type") => &mut attrs.value_type,
            (OFFICE, "value") => &mut attrs.value,
            (OFFICE, "string-value") => &mut attrs.string_value,
            (OFFICE, "boolean-value") => &mut attrs.boolean_value,
            (OFFICE, "date-value") => &mut attrs.date_value,
            (TABLE, "date-value") => &mut attrs.null_date,
            (OFFICE, "time-value") => &mut attrs.time_value,
            (CALCEXT, "value-type") => &mut attrs.calcext_type,
            (TABLE, "formula") => &mut attrs.formula,
            (TABLE, "cell-range-address") => &mut attrs.range_address,
            (TABLE, "expression") => &mut attrs.expression,
            (TABLE, "number-columns-repeated") => &mut attrs.columns_repeated,
            (TABLE, "number-rows-repeated") => &mut attrs.rows_repeated,
            (TABLE, "number-matrix-columns-spanned") => &mut attrs.matrix_columns,
            (TABLE, "number-matrix-rows-spanned") => &mut attrs.matrix_rows,
            (TABLE | STYLE, "name") => &mut attrs.name,
            (TABLE, "style-name") => &mut attrs.style_name,
            (TABLE, "display") => &mut attrs.display,
            (STYLE, "family") => &mut attrs.family,
            (TEXT, "c") => &mut attrs.count,
            _ => continue,
        };
        *slot = Some(value);
    }
    Ok(attrs)
}

fn classify(resolved: &ResolveResult<'_>, local: &str) -> Tag {
    let ResolveResult::Bound(namespace) = resolved else {
        return Tag::Other;
    };
    match (namespace.0, local) {
        (TABLE, "table") => Tag::Table,
        (TABLE, "table-row") => Tag::Row,
        (TABLE, "table-row-group" | "table-header-rows" | "table-rows") => Tag::RowContainer,
        (TABLE, "table-cell") => Tag::Cell,
        (TABLE, "covered-table-cell") => Tag::CoveredCell,
        (TABLE, "null-date") => Tag::NullDate,
        (TABLE, "named-expressions") => Tag::NamedExpressions,
        (TABLE, "named-range") => Tag::NamedRange,
        (TABLE, "named-expression") => Tag::NamedExpression,
        (TEXT, "p") => Tag::Paragraph,
        (TEXT, "s") => Tag::Space,
        (TEXT, "tab") => Tag::Tab,
        (TEXT, "line-break") => Tag::LineBreak,
        (OFFICE, "annotation") | (TEXT, "note" | "ruby-text") => Tag::Aside,
        (TEXT, "span") => Tag::Span,
        (TEXT, "a") => Tag::Link,
        (STYLE, "style") => Tag::Style,
        (STYLE, "table-properties") => Tag::TableProperties,
        _ => Tag::Other,
    }
}

fn reference_text(reference: &BytesRef<'_>, part: &str) -> Result<String, Error> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|source| Error::xml(part, source))?
    {
        return Ok(character.to_string());
    }
    let name = reference.xml10_content();
    resolve_xml_entity(&name)
        .map(str::to_owned)
        .ok_or_else(|| Error::malformed(part, format!("unknown entity &{name};")))
}
