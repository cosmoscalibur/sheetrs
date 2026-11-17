//! A lexer of XML for the parts of a workbook that have millions of elements: the sheets.
//!
//! [`Lexer`] gives the markup and the text of a document as slices of the window it keeps over the
//! source, one token at a time, and nothing else: no event is built, the attributes are not
//! parsed until a caller asks for them, and the text is not unescaped. A reader that wants the
//! events of `quick-xml` for the rare element (a merged range, a column) builds them from a
//! [`Tag`] ([`Tag::to_start`]); one that wants the value of a cell reads the bytes it needs.
//!
//! What it checks is what a reader of a workbook relies on: the bytes are UTF-8, the markup is
//! closed, and an end tag has the name of the element it closes. It does not know namespaces or
//! encodings other than UTF-8, and it does not expand entities.
//!
//! It exists because the tokenizer and the events of `quick-xml` were about 45 % of the time of
//! reading a sheet (see `docs/performance.md`), and a reader of events cannot skip the part of an
//! event that a reader of cells never looks at: a reader that builds the same events over a faster
//! tokenizer was measured at the same speed as `quick-xml`.

use std::borrow::Cow;
use std::io::{self, Read};
use std::ops::Range;

use memchr::memchr;
use quick_xml::XmlVersion;
use quick_xml::errors::{IllFormedError, SyntaxError};
use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesRef, BytesStart};
use quick_xml::name::QName;
use quick_xml::parser::{ElementParser, Parser, PiParser};

pub use quick_xml::{Error, Result};

/// The size of the window over the source. A token longer than this makes it grow.
const WINDOW: usize = 128 * 1024;

const fn is_whitespace(b: u8) -> bool {
    matches!(b, b' ' | b'\r' | b'\n' | b'\t')
}

/// A start or an empty tag: its name and its attributes, as they are in the document.
#[derive(Clone, Copy, Debug)]
pub struct Tag<'a> {
    raw: &'a [u8],
    name_len: usize,
}

impl<'a> Tag<'a> {
    /// The name, with its prefix if it has one.
    #[must_use]
    pub fn name(&self) -> &'a [u8] {
        &self.raw[..self.name_len]
    }

    /// The attributes, in the order of the document. A malformed attribute ends the iteration.
    #[must_use]
    pub fn attributes(&self) -> Attrs<'a> {
        Attrs {
            rest: &self.raw[self.name_len..],
        }
    }

    /// The tag as the event of `quick-xml`, for the code that reads attributes with it.
    ///
    /// # Errors
    ///
    /// When the tag is not UTF-8 (a lexer has checked it already).
    pub fn to_start(&self) -> Result<BytesStart<'a>> {
        Ok(BytesStart::from_content(
            std::str::from_utf8(self.raw)?,
            self.name_len,
        ))
    }
}

/// The attributes of a tag, as pairs of a name and a raw value (what is between the quotes).
#[derive(Clone, Debug)]
pub struct Attrs<'a> {
    rest: &'a [u8],
}

impl<'a> Attrs<'a> {
    /// The attributes in `raw`: what follows the name in a tag.
    #[must_use]
    pub fn new(raw: &'a [u8]) -> Self {
        Self { rest: raw }
    }
}

impl<'a> Iterator for Attrs<'a> {
    type Item = (&'a [u8], &'a [u8]);

    fn next(&mut self) -> Option<Self::Item> {
        let rest = self.rest.trim_ascii_start();
        let eq = memchr(b'=', rest)?;
        let key = rest[..eq].trim_ascii_end();
        let after = rest[eq + 1..].trim_ascii_start();
        let quote = *after.first()?;
        let len = if quote == b'"' || quote == b'\'' {
            memchr(quote, &after[1..])
        } else {
            None
        };
        let Some(len) = len else {
            self.rest = &[];
            return None;
        };
        self.rest = &after[len + 2..];
        Some((key, &after[1..=len]))
    }
}

/// The value of an attribute as `quick-xml` gives it: entities resolved and the white space of the
/// document (tab, line break) read as spaces, as XML 1.0 says.
///
/// # Errors
///
/// When the value is not UTF-8 or holds an entity that cannot be resolved.
pub fn attr_value(raw: &[u8]) -> Result<Cow<'_, str>> {
    let text = std::str::from_utf8(raw)?;
    if !raw
        .iter()
        .any(|&b| matches!(b, b'&' | b'\t' | b'\r' | b'\n'))
    {
        return Ok(Cow::Borrowed(text));
    }
    Attribute {
        key: QName(""),
        value: Cow::Borrowed(text),
    }
    .normalized_value(XmlVersion::Implicit1_0)
}

/// Append the text of `raw` to `out`, with each reference (`&amp;`, `&#10;`) given to `resolve`,
/// which appends what it stands for.
///
/// # Errors
///
/// When the text is not UTF-8, a reference is not closed, or `resolve` fails.
pub fn unescape_into<E: From<Error>>(
    raw: &[u8],
    out: &mut String,
    mut resolve: impl FnMut(&mut String, &BytesRef<'_>) -> std::result::Result<(), E>,
) -> std::result::Result<(), E> {
    let text = std::str::from_utf8(raw).map_err(Error::from)?;
    let mut rest = text;
    while let Some(amp) = memchr(b'&', rest.as_bytes()) {
        out.push_str(&rest[..amp]);
        let after = &rest[amp + 1..];
        // The reference ends at the `;` that comes before another `&`.
        let end = memchr::memchr2(b';', b'&', after.as_bytes())
            .filter(|&i| after.as_bytes()[i] == b';')
            .ok_or(Error::IllFormed(IllFormedError::UnclosedReference))?;
        resolve(out, &BytesRef::new(&after[..end]))?;
        rest = &after[end + 1..];
    }
    out.push_str(rest);
    Ok(())
}

/// What the lexer gives.
#[derive(Clone, Copy, Debug)]
pub enum Token<'a> {
    /// `<name ...>`
    Start(Tag<'a>),
    /// `<name .../>`
    Empty(Tag<'a>),
    /// `</name>`: the name.
    End(&'a [u8]),
    /// Text between markup, never empty, as it is in the document (references included).
    Text(&'a [u8]),
    /// The content of `<![CDATA[...]]>`.
    CData(&'a [u8]),
    /// A comment, a declaration, a processing instruction or a doctype.
    Other,
    /// The end of the document.
    Eof,
}

/// Where a token is in the window.
enum Span {
    Start {
        content: Range<usize>,
        name_len: usize,
    },
    Empty {
        content: Range<usize>,
        name_len: usize,
    },
    End(Range<usize>),
    Text(Range<usize>),
    CData(Range<usize>),
    Other,
    Eof,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Markup {
    Tag,
    Pi,
    Comment,
    CData,
    DocType,
}

impl Markup {
    /// The index of the `>` that closes the markup in `bytes` (what follows the `<`).
    fn end(self, bytes: &[u8]) -> Option<usize> {
        match self {
            Markup::Tag => ElementParser::Outside.feed(bytes),
            Markup::Pi => PiParser(false).feed(bytes),
            Markup::Comment => {
                // `<!---->` is the shortest: the `>` is at 5 or more after `<`.
                let mut from = 0;
                while let Some(i) = memchr(b'>', &bytes[from..]) {
                    let at = from + i;
                    if at >= 5 && bytes[..at].ends_with(b"--") {
                        return Some(at);
                    }
                    from = at + 1;
                }
                None
            }
            Markup::CData => {
                let mut from = 0;
                while let Some(i) = memchr(b'>', &bytes[from..]) {
                    let at = from + i;
                    if bytes[..at].ends_with(b"]]") {
                        return Some(at);
                    }
                    from = at + 1;
                }
                None
            }
            Markup::DocType => doctype_end(bytes),
        }
    }

    fn unclosed(self, bytes: &[u8]) -> SyntaxError {
        match self {
            Markup::Tag => {
                let mut parser = ElementParser::Outside;
                parser.feed(bytes);
                parser.eof_error(bytes)
            }
            Markup::Pi => PiParser(false).eof_error(bytes),
            Markup::Comment => SyntaxError::UnclosedComment,
            Markup::CData => SyntaxError::UnclosedCData,
            Markup::DocType => SyntaxError::UnclosedDoctype,
        }
    }
}

/// The `>` that ends a `<!DOCTYPE`, outside of quotes, comments and the internal subset (`[...]`).
fn doctype_end(bytes: &[u8]) -> Option<usize> {
    let (mut quote, mut subset, mut i) = (0u8, false, 0);
    while i < bytes.len() {
        let b = bytes[i];
        if quote != 0 {
            if b == quote {
                quote = 0;
            }
        } else if subset && bytes[i..].starts_with(b"<!--") {
            // A comment of the subset may hold a `]` or a `>`.
            let end = bytes[i + 4..].windows(3).position(|w| w == b"-->")?;
            i += 4 + end + 2;
        } else {
            match b {
                b'"' | b'\'' => quote = b,
                b'[' => subset = true,
                b']' => subset = false,
                b'>' if !subset => return Some(i),
                _ => {}
            }
        }
        i += 1;
    }
    None
}

/// A lexer over a [`Read`].
#[derive(Debug)]
pub struct Lexer<R> {
    source: R,
    /// The window: `buf[pos..end]` is what has been read and not given, `buf[..valid]` is UTF-8.
    buf: Vec<u8>,
    pos: usize,
    end: usize,
    valid: usize,
    eof: bool,
    started: bool,
    /// The bytes given before the window was last moved.
    base: u64,
    /// The length of the byte order mark, which is not counted in the position.
    bom: u64,
    /// The names of the elements that are open, one after the other, and where each starts.
    names: Vec<u8>,
    starts: Vec<usize>,
}

impl<R: Read> Lexer<R> {
    /// A lexer of the bytes of `source`.
    pub fn new(source: R) -> Self {
        Self {
            source,
            buf: vec![0; WINDOW],
            pos: 0,
            end: 0,
            valid: 0,
            eof: false,
            started: false,
            base: 0,
            bom: 0,
            names: Vec::new(),
            starts: Vec::new(),
        }
    }

    /// The bytes given so far, a byte order mark apart: where the next token starts.
    #[must_use]
    pub fn position(&self) -> u64 {
        self.base + self.pos as u64 - self.bom
    }

    /// Read more into the window, keeping what has not been given. Whether it holds more than before.
    fn fill(&mut self) -> Result<bool> {
        if self.eof {
            return Ok(false);
        }
        if self.pos > 0 {
            self.buf.copy_within(self.pos..self.end, 0);
            self.base += self.pos as u64;
            self.end -= self.pos;
            self.valid = self.valid.saturating_sub(self.pos);
            self.pos = 0;
        }
        if self.end == self.buf.len() {
            let len = self.buf.len() * 2;
            self.buf.resize(len, 0);
        }
        let n = loop {
            match self.source.read(&mut self.buf[self.end..]) {
                Err(e) if e.kind() == io::ErrorKind::Interrupted => {}
                Err(e) => return Err(e.into()),
                Ok(n) => break n,
            }
        };
        self.end += n;
        if n == 0 {
            self.eof = true;
        }
        match std::str::from_utf8(&self.buf[self.valid..self.end]) {
            Ok(_) => self.valid = self.end,
            Err(e) if e.error_len().is_none() && !self.eof => self.valid += e.valid_up_to(),
            Err(e) => return Err(e.into()),
        }
        Ok(n > 0)
    }

    /// The next token. It borrows from the window: a caller keeps what it needs before the next call.
    ///
    /// # Errors
    ///
    /// When the source cannot be read or is not UTF-8, when markup is not closed, or an end tag does
    /// not close the element that is open.
    pub fn token(&mut self) -> Result<Token<'_>> {
        let (span, _) = self.scan()?;
        Ok(match span {
            Span::Start { content, name_len } => Token::Start(Tag {
                raw: &self.buf[content],
                name_len,
            }),
            Span::Empty { content, name_len } => Token::Empty(Tag {
                raw: &self.buf[content],
                name_len,
            }),
            Span::End(name) => Token::End(&self.buf[name]),
            Span::Text(text) => Token::Text(&self.buf[text]),
            Span::CData(text) => Token::CData(&self.buf[text]),
            Span::Other => Token::Other,
            Span::Eof => Token::Eof,
        })
    }

    /// Where the next token is in the window, and the bytes it takes.
    fn scan(&mut self) -> Result<(Span, Range<usize>)> {
        if !self.started {
            self.started = true;
            while self.end < 3 && self.fill()? {}
            if self.buf[..self.end].starts_with(b"\xEF\xBB\xBF") {
                self.pos = 3;
                self.bom = 3;
            }
        }
        if self.pos == self.end && !self.fill()? {
            return Ok((Span::Eof, 0..0));
        }
        if self.buf[self.pos] == b'<' {
            self.markup()
        } else {
            self.text()
        }
    }

    fn text(&mut self) -> Result<(Span, Range<usize>)> {
        // Look for the `<` from where the last look ended, over the bytes that arrive.
        let mut seen = 0;
        let len = loop {
            if let Some(i) = memchr(b'<', &self.buf[self.pos + seen..self.end]) {
                break seen + i;
            }
            seen = self.end - self.pos;
            if !self.fill()? {
                break self.end - self.pos;
            }
        };
        let start = self.pos;
        self.pos += len;
        Ok((Span::Text(start..start + len), start..start + len))
    }

    fn markup(&mut self) -> Result<(Span, Range<usize>)> {
        // The kind of markup needs the two bytes after the `<`.
        while self.end - self.pos < 3 && self.fill()? {}
        let after = &self.buf[self.pos + 1..self.end];
        let kind = match after.first() {
            None => return Err(Error::Syntax(SyntaxError::UnclosedTag)),
            Some(b'?') => Markup::Pi,
            Some(b'!') => match after.get(1) {
                Some(b'[') => Markup::CData,
                Some(b'-') => Markup::Comment,
                Some(b'D' | b'd') => Markup::DocType,
                _ => return Err(Error::Syntax(SyntaxError::InvalidBangMarkup)),
            },
            Some(_) => Markup::Tag,
        };
        // Up to the `>`: the window grows until the markup is whole in it.
        let close = loop {
            if let Some(i) = kind.end(&self.buf[self.pos + 1..self.end]) {
                break self.pos + 1 + i;
            }
            if !self.fill()? {
                return Err(Error::Syntax(
                    kind.unclosed(&self.buf[self.pos + 1..self.end]),
                ));
            }
        };
        let start = self.pos;
        self.pos = close + 1;
        let raw = start..close + 1;
        let token = &self.buf[start..=close];
        match kind {
            Markup::Tag if token.get(1) == Some(&b'/') => {
                // Space after the name is allowed.
                let name = token[2..token.len() - 1].trim_ascii_end();
                let Some(open) = self.starts.pop() else {
                    return Err(Error::IllFormed(IllFormedError::UnmatchedEndTag(
                        String::from_utf8_lossy(name).into_owned(),
                    )));
                };
                if self.names[open..] != *name {
                    let expected = String::from_utf8_lossy(&self.names[open..]).into_owned();
                    self.names.truncate(open);
                    return Err(Error::IllFormed(IllFormedError::MismatchedEndTag {
                        expected,
                        found: String::from_utf8_lossy(name).into_owned(),
                    }));
                }
                self.names.truncate(open);
                Ok((Span::End(start + 2..start + 2 + name.len()), raw))
            }
            Markup::Tag => {
                let inner = &token[1..token.len() - 1];
                let (content, empty) = match inner.strip_suffix(b"/") {
                    Some(content) => (content, true),
                    None => (inner, false),
                };
                let name_len = content
                    .iter()
                    .position(|&b| is_whitespace(b))
                    .unwrap_or(content.len());
                if !empty {
                    self.starts.push(self.names.len());
                    self.names.extend_from_slice(&content[..name_len]);
                }
                let content = start + 1..start + 1 + content.len();
                Ok((
                    if empty {
                        Span::Empty { content, name_len }
                    } else {
                        Span::Start { content, name_len }
                    },
                    raw,
                ))
            }
            Markup::CData if token.starts_with(b"<![CDATA[") => {
                Ok((Span::CData(start + 9..start + token.len() - 3), raw))
            }
            Markup::Pi | Markup::Comment | Markup::DocType => Ok((Span::Other, raw)),
            Markup::CData => Err(Error::Syntax(kind.unclosed(token))),
        }
    }

    /// Give up to the end tag of the element that has just been opened, everything inside it
    /// included.
    ///
    /// # Errors
    ///
    /// When the document is not well formed or ends before the end tag.
    pub fn skip_element(&mut self) -> Result<()> {
        let mut depth = 0usize;
        loop {
            match self.token()? {
                Token::Start(_) => depth += 1,
                Token::End(_) => {
                    if depth == 0 {
                        return Ok(());
                    }
                    depth -= 1;
                }
                Token::Eof => {
                    return Err(Error::IllFormed(IllFormedError::MissingEndTag(
                        String::new(),
                    )));
                }
                _ => {}
            }
        }
    }
}
