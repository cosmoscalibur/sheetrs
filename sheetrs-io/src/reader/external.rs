//! Where the files that a workbook's formulas refer to are, as `LibreOffice` resolves them: against
//! the folder of the document, and written as a `file://` URL.

use super::Format;
use std::path::{Component, Path, PathBuf};

/// An external workbook of a formula, and where it is.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalUrl {
    /// The file name (`test.xlsx`).
    pub basename: String,
    /// The URL `LibreOffice` shows (`file:///folder/test.xlsx`).
    pub url: String,
    /// How the formulas of the file write it.
    pub address: String,
}

/// `%20` is a space, as in the URL of a file.
fn percent_decoded(text: &str) -> String {
    let bytes = text.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%'
            && i + 2 < bytes.len()
            && let Ok(byte) = u8::from_str_radix(&text[i + 1..i + 3], 16)
        {
            out.push(byte);
            i += 3;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    String::from_utf8_lossy(&out).into_owned()
}

/// The path with `.` and `..` taken out, with no look at the disk.
fn normalized(path: &Path) -> PathBuf {
    let mut out = PathBuf::new();
    for part in path.components() {
        match part {
            Component::CurDir => {}
            Component::ParentDir => {
                if !out.pop() {
                    out.push("..");
                }
            }
            other => out.push(other),
        }
    }
    out
}

/// The URL of the external workbook that a file names with `target` (the `Target` of its
/// relationship in an XLSX, the `xlink:href` of its cached table in an ODS), as `LibreOffice`
/// shows it in a formula: an absolute location stays, a relative one is resolved against the
/// folder of `document` (the document itself is the folder in an ODS, whose links are relative to
/// the package: `../test.xlsx` is the file next to the document). A URL of another scheme, or any
/// target when the document has no location (its bytes), is returned as it is.
#[must_use]
pub fn external_url(format: Format, document: Option<&Path>, target: &str) -> String {
    let target = percent_decoded(target);
    if let Some(path) = target.strip_prefix("file://") {
        return format!("file://{}", normalized(Path::new(path)).display());
    }
    if target.contains("://") {
        return target;
    }
    if target.starts_with('/') {
        return format!("file://{}", normalized(Path::new(&target)).display());
    }
    let Some(document) = document.and_then(|d| std::path::absolute(d).ok()) else {
        return target;
    };
    let base = match format {
        Format::Ods => document,
        _ => document.parent().map(Path::to_path_buf).unwrap_or_default(),
    };
    format!("file://{}", normalized(&base.join(&target)).display())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_target_is_resolved_as_libreoffice_resolves_it() {
        let doc = Some(Path::new("/work/books/doc.xlsx"));
        let url = |f, t| external_url(f, doc, t);
        assert_eq!(
            url(Format::Xlsx, "test.xlsx"),
            "file:///work/books/test.xlsx"
        );
        assert_eq!(
            url(Format::Xlsx, "./sub/test.xlsx"),
            "file:///work/books/sub/test.xlsx"
        );
        assert_eq!(
            url(Format::Xlsx, "../other/test.xlsx"),
            "file:///work/other/test.xlsx"
        );
        assert_eq!(
            url(Format::Xlsx, "file:///some/dir/test.xlsx"),
            "file:///some/dir/test.xlsx"
        );
        assert_eq!(
            url(Format::Xlsx, "/some/dir/test.xlsx"),
            "file:///some/dir/test.xlsx"
        );
        assert_eq!(
            url(Format::Xlsx, "file:///some/my%20dir/test%20two.xlsx"),
            "file:///some/my dir/test two.xlsx"
        );
        let ods = Some(Path::new("/work/books/doc.ods"));
        assert_eq!(
            external_url(Format::Ods, ods, "../test.xlsx"),
            "file:///work/books/test.xlsx"
        );
        assert_eq!(external_url(Format::Xlsx, None, "test.xlsx"), "test.xlsx");
    }
}
