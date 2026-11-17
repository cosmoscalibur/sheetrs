//! The lexer of `sheetrs_io::xml` sees what `quick_xml::Reader` sees: the same elements with the
//! same attributes, the same text (references and all), the same CDATA, in the same order, at the
//! same positions, and the same documents are refused.

use std::fs;
use std::io::{BufReader, Read};
use std::path::{Path, PathBuf};

use quick_xml::Error;
use quick_xml::events::Event;
use sheetrs_io::xml::{Lexer, Token, attr_value, unescape_into};

/// A source that gives its bytes a few at a time, so that a token is split at every place.
struct Chunked<'a>(&'a [u8], usize);

impl Read for Chunked<'_> {
    fn read(&mut self, out: &mut [u8]) -> std::io::Result<usize> {
        let n = self.0.len().min(self.1).min(out.len());
        out[..n].copy_from_slice(&self.0[..n]);
        self.0 = &self.0[n..];
        Ok(n)
    }
}

/// What a document is made of, as text: the same for both readers.
type Stream = Vec<String>;

fn quick_xml_stream(bytes: &[u8]) -> (Stream, Option<Error>) {
    let mut reader = quick_xml::Reader::from_reader(BufReader::new(bytes));
    let (mut out, mut buf, mut text) = (Stream::new(), Vec::new(), String::new());
    let flush = |out: &mut Stream, text: &mut String| {
        if !text.is_empty() {
            out.push(format!("text {text:?}"));
            text.clear();
        }
    };
    loop {
        buf.clear();
        let step = reader.read_event_into(&mut buf);
        let at = reader.buffer_position();
        match step {
            Err(e) => {
                flush(&mut out, &mut text);
                return (out, Some(e));
            }
            Ok(Event::Text(t)) => text.push_str(&t),
            Ok(Event::GeneralRef(r)) => {
                text.push('&');
                text.push_str(&r);
                text.push(';');
            }
            Ok(event) => {
                flush(&mut out, &mut text);
                match event {
                    Event::Start(e) => out.push(format!(
                        "start {} [{}] @{at}",
                        e.name().as_ref(),
                        e.attributes_raw()
                    )),
                    Event::Empty(e) => out.push(format!(
                        "empty {} [{}] @{at}",
                        e.name().as_ref(),
                        e.attributes_raw()
                    )),
                    Event::End(e) => out.push(format!("end {} @{at}", e.name().as_ref())),
                    Event::CData(c) => out.push(format!("cdata {:?}", &*c)),
                    Event::Eof => return (out, None),
                    _ => out.push("other".to_string()),
                }
            }
        }
    }
}

/// The streams of both, with the start and empty tags told apart the same way.
fn canonical_lexer(bytes: &[u8], chunk: usize) -> (Stream, Option<Error>) {
    let mut lexer = Lexer::new(Chunked(bytes, chunk));
    let mut out = Stream::new();
    loop {
        // The token borrows the lexer: it becomes text before the position is asked.
        let (entry, positioned) = match lexer.token() {
            Err(e) => return (out, Some(e)),
            Ok(Token::Eof) => return (out, None),
            Ok(Token::Text(t)) => (format!("text {:?}", String::from_utf8_lossy(t)), false),
            Ok(Token::CData(c)) => (format!("cdata {:?}", String::from_utf8_lossy(c)), false),
            Ok(Token::Other) => ("other".to_string(), false),
            Ok(Token::Start(t)) => {
                let s = t.to_start().unwrap();
                (
                    format!("start {} [{}]", s.name().as_ref(), s.attributes_raw()),
                    true,
                )
            }
            Ok(Token::Empty(t)) => {
                let s = t.to_start().unwrap();
                (
                    format!("empty {} [{}]", s.name().as_ref(), s.attributes_raw()),
                    true,
                )
            }
            Ok(Token::End(name)) => (format!("end {}", String::from_utf8_lossy(name)), true),
        };
        out.push(if positioned {
            format!("{entry} @{}", lexer.position())
        } else {
            entry
        });
    }
}

/// Whether both refuse `bytes` alike, or accept it and see the same.
fn assert_same(bytes: &[u8]) {
    let (expected, expected_error) = quick_xml_stream(bytes);
    // A dangling `&` is for whoever unescapes the text: the lexer does not look inside it.
    if matches!(
        expected_error,
        Some(Error::IllFormed(
            quick_xml::errors::IllFormedError::UnclosedReference
        ))
    ) {
        return;
    }
    for chunk in [1, 2, 3, 7, 64, usize::MAX] {
        let (got, error) = canonical_lexer(bytes, chunk);
        assert_eq!(
            (got, error.map(|e| e.to_string())),
            (
                expected.clone(),
                expected_error.as_ref().map(ToString::to_string)
            ),
            "chunks of {chunk}: {:?}",
            String::from_utf8_lossy(bytes)
        );
    }
}

#[test]
fn markup_of_every_kind_is_seen_alike() {
    for text in [
        "",
        "text only",
        "<a/>",
        "<a>text</a>",
        "<a x=\"1\" y='2'>t</a>",
        "<a x=\">\" y='>'>t</a>",
        "<a>1 &amp; 2 &lt; 3 &#10; &#x41;</a>",
        "<a>  spaced  &amp;  out  </a>",
        "  <a>  </a>  ",
        "\u{feff}<a>bom</a>",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<a/>",
        "<?pi content?><a/>",
        "<a><!-- a comment --></a>",
        "<a><!----></a>",
        "<a><!-- a > b --></a>",
        "<a><![CDATA[ x > y ]]></a>",
        "<a><![CDATA[]]></a>",
        "<!DOCTYPE a><a/>",
        "<!DOCTYPE a [ <!ENTITY e \"v\"> ]><a/>",
        "<a>é ü 日本</a>",
        "<a b=\"é\"/>",
        "<a>\r\n</a>",
        "<a></b>",
        "</a>",
        "<a>",
        "<a",
        "<a x=\"1",
        "<!--",
        "<![CDATA[x",
        "<!x>",
        "<!",
        "<?",
        "<a><</a>",
        "<a>>\"'</a>",
        "<a/><b/>",
        "<a></a >",
        "<a  x = \"1\"  />",
        "<a><b><c/></b></a>",
        "<a>1</a><a>2</a>",
    ] {
        assert_same(text.as_bytes());
    }
}

#[test]
fn invalid_utf8_is_refused_alike() {
    for bytes in [
        &b"<a>\xff</a>"[..],
        b"<a \xff=\"1\"/>",
        b"<a>ok\xc3</a>",
        b"\xff",
    ] {
        let (_, theirs) = quick_xml_stream(bytes);
        for chunk in [1, 2, usize::MAX] {
            let (_, ours) = canonical_lexer(bytes, chunk);
            assert!(theirs.is_some() && ours.is_some(), "{bytes:?}");
        }
    }
}

#[test]
fn random_markup_is_seen_alike() {
    let pieces = [
        "<a>",
        "</a>",
        "<b x=\"1\" y='2'>",
        "</b>",
        "<c/>",
        "<c x=\"&amp;\"/>",
        "text",
        " ",
        "  ",
        "\n",
        "&amp;",
        ";",
        "&#10;",
        "<!-- c -->",
        "<!--",
        "-->",
        "<![CDATA[x]]>",
        "<![CDATA[",
        "]]>",
        "<?xml version=\"1.0\"?>",
        "<?pi x?>",
        "<!DOCTYPE a>",
        "é",
        "<",
        ">",
        "\"",
        "'",
        "<a b=\">\">",
        "</a >",
        "<!-->",
        "<!--->",
        "<d",
        "/",
        "=",
        "?",
        "!",
        "-",
        "]",
    ];
    let mut seed = 0x2545_f491_4f6c_dd1d_u64;
    for _ in 0..4000 {
        let mut text = String::new();
        seed ^= seed << 13;
        seed ^= seed >> 7;
        seed ^= seed << 17;
        let count = 1 + (seed % 12) as usize;
        for _ in 0..count {
            seed ^= seed << 13;
            seed ^= seed >> 7;
            seed ^= seed << 17;
            text.push_str(pieces[(seed % pieces.len() as u64) as usize]);
        }
        assert_same(text.as_bytes());
    }
}

/// The workbooks of the repository.
fn workbooks() -> Vec<PathBuf> {
    fn walk(dir: &Path, out: &mut Vec<PathBuf>) {
        let Ok(entries) = fs::read_dir(dir) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                if path
                    .file_name()
                    .is_some_and(|n| n != "target" && n != ".git")
                {
                    walk(&path, out);
                }
            } else if matches!(
                path.extension().and_then(|e| e.to_str()),
                Some("xlsx" | "ods" | "xlsm")
            ) {
                out.push(path);
            }
        }
    }
    let mut out = Vec::new();
    walk(&Path::new(env!("CARGO_MANIFEST_DIR")).join(".."), &mut out);
    out.sort();
    out
}

#[test]
fn the_parts_of_every_workbook_of_the_repository_are_seen_alike() {
    let books = workbooks();
    assert!(books.len() > 50, "found only {} workbooks", books.len());
    let mut parts = 0;
    for path in books {
        let Ok(file) = fs::File::open(&path) else {
            continue;
        };
        let Ok(mut archive) = zip::ZipArchive::new(file) else {
            continue;
        };
        for i in 0..archive.len() {
            let Ok(mut entry) = archive.by_index(i) else {
                continue;
            };
            let name = entry
                .name()
                .map(std::borrow::Cow::into_owned)
                .unwrap_or_default();
            let is_xml = Path::new(&name)
                .extension()
                .is_some_and(|e| e.eq_ignore_ascii_case("xml") || e.eq_ignore_ascii_case("rels"));
            if !is_xml {
                continue;
            }
            let mut bytes = Vec::new();
            if entry.read_to_end(&mut bytes).is_err() {
                continue;
            }
            let (expected, expected_error) = quick_xml_stream(&bytes);
            let (got, error) = canonical_lexer(&bytes, 4096);
            assert!(error.is_none() && expected_error.is_none(), "{name}");
            assert!(
                got == expected,
                "{} / {name}: first difference at token {}",
                path.display(),
                got.iter()
                    .zip(&expected)
                    .position(|(a, b)| a != b)
                    .unwrap_or(0)
            );
            parts += 1;
        }
    }
    assert!(parts > 200, "compared only {parts} parts");
}

#[test]
fn attribute_values_are_the_ones_of_quick_xml() {
    for value in [
        "plain",
        "a &amp; b",
        "tab\there",
        "line\nbreak",
        "cr\r\nlf",
        "&#10;kept",
        "&lt;&gt;&quot;&apos;",
        "&#x41;&#66;",
        "é",
    ] {
        let tag = format!("<a v=\"{value}\"/>");
        let mut reader = quick_xml::Reader::from_str(&tag);
        let Ok(Event::Empty(e)) = reader.read_event() else {
            panic!("{tag}");
        };
        let theirs = e
            .attributes()
            .next()
            .unwrap()
            .unwrap()
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap()
            .into_owned();
        let mut lexer = Lexer::new(tag.as_bytes());
        let Token::Empty(t) = lexer.token().unwrap() else {
            panic!("{tag}");
        };
        let (key, raw) = t.attributes().next().unwrap();
        assert_eq!(key, b"v");
        assert_eq!(attr_value(raw).unwrap(), theirs, "{value:?}");
    }
}

#[test]
fn unescaped_text_has_what_its_references_stand_for() {
    let resolve = |out: &mut String, r: &quick_xml::events::BytesRef<'_>| -> Result<(), Error> {
        if let Some(c) = r.resolve_char_ref()? {
            out.push(c);
        } else if let Some(s) = quick_xml::escape::resolve_predefined_entity(r) {
            out.push_str(s);
        } else {
            out.push('&');
            out.push_str(r);
            out.push(';');
        }
        Ok(())
    };
    let mut out = String::new();
    unescape_into(b"a &amp; b &lt;&#10;&#x41; &nbsp; c", &mut out, resolve).unwrap();
    assert_eq!(out, "a & b <\nA &nbsp; c");
    assert!(unescape_into(b"a & b", &mut String::new(), resolve).is_err());
    assert!(unescape_into(b"a &amp", &mut String::new(), resolve).is_err());
}
