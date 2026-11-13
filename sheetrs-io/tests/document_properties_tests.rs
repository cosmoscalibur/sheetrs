//! The document properties that the readers give for each file are named `created`, `modified` and
//! `generator` (a decision of `SheetRS`: the files name them in their own way), whichever the format.

use sheetrs_io::Book;

fn properties(extension: &str) -> Vec<(String, String)> {
    let path = format!(
        "{}/../tests/lint_book.{extension}",
        env!("CARGO_MANIFEST_DIR")
    );
    Book::from_file(path)
        .unwrap()
        .workbook()
        .unwrap()
        .properties
}

#[test]
fn each_file_gives_its_properties_under_the_names_of_sheetrs() {
    for extension in ["xlsx", "ods"] {
        let properties = properties(extension);
        for key in ["created", "modified", "generator"] {
            assert!(
                properties.iter().any(|(k, _)| k == key),
                "{extension} {key}"
            );
        }
    }
}

/// The last-modified date of each book is 2020-01-15T10:00:00. The `dc:date` of the ODS has no time
/// zone and a fraction of a second, as `LibreOffice` writes it: it is read as UTC.
#[test]
fn the_modified_date_is_read_with_or_without_a_time_zone() {
    for extension in ["xlsx", "ods"] {
        let path = format!(
            "{}/../tests/lint_book.{extension}",
            env!("CARGO_MANIFEST_DIR")
        );
        let modified = Book::from_file(path)
            .unwrap()
            .workbook()
            .unwrap()
            .modified_date
            .unwrap();
        assert_eq!(
            modified.to_rfc3339(),
            "2020-01-15T10:00:00+00:00",
            "{extension}"
        );
    }
}
