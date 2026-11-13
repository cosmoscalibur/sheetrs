//! The scope of the defined names of `tests/lint_book.{xlsx,ods}`. `LibreOffice` writes the names of
//! the document of an ODS after its last table, which in this book is the cache of an external
//! workbook: the scope of those names is the workbook, not that table.

use sheetrs_io::Book;
use sheetrs_io::reader::NameScope;

#[test]
fn a_name_defined_twice_has_the_workbook_and_the_sheet_as_scopes_in_each_format() {
    for extension in ["xlsx", "ods"] {
        let path = format!(
            "{}/../tests/lint_book.{extension}",
            env!("CARGO_MANIFEST_DIR")
        );
        let workbook = Book::from_file(path).unwrap().workbook().unwrap();
        let conflict = workbook
            .name_conflicts
            .iter()
            .find(|c| c.name == "Local")
            .unwrap_or_else(|| panic!("{extension}: no conflict for `Local`"));
        assert_eq!(
            conflict.scopes,
            [NameScope::Workbook, NameScope::Sheet("Calc".to_string())],
            "{extension}"
        );
        // No sheet is named as the address of the external workbook cache.
        assert!(
            conflict.scopes.iter().all(|s| match s {
                NameScope::Sheet(name) => !name.starts_with('\''),
                NameScope::Workbook => true,
            }),
            "{extension}"
        );
    }
}
