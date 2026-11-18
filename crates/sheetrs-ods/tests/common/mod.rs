use std::error::Error as StdError;
use std::io::{Cursor, Write};

use zip::CompressionMethod;
use zip::write::{SimpleFileOptions, ZipWriter};

const NAMESPACES: &str = r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2""#;

/// The `content.xml` of a one-sheet document named `S1` whose table holds `rows`.
pub fn content_xml(rows: &str, settings: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content {NAMESPACES} office:version="1.3"><office:body><office:spreadsheet>{settings}<table:table table:name="S1">{rows}</table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

/// An ODS package whose `content.xml` is `content`.
pub fn package(content: &str) -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    writer.start_file("mimetype", options)?;
    writer.write_all(b"application/vnd.oasis.opendocument.spreadsheet")?;
    writer.start_file("content.xml", options)?;
    writer.write_all(content.as_bytes())?;
    writer.start_file("META-INF/manifest.xml", options)?;
    writer.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
    )?;
    Ok(writer.finish()?)
}

/// A `content.xml` with `styles` in its automatic styles and `tables` as the tables of the
/// spreadsheet.
pub fn document(styles: &str, tables: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content {NAMESPACES} office:version="1.3"><office:automatic-styles>{styles}</office:automatic-styles><office:body><office:spreadsheet>{tables}</office:spreadsheet></office:body></office:document-content>"#
    )
}
