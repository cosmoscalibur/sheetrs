//! The empty search text matches nothing, in every mode, as in `LibreOffice` (`findAll` reports no
//! cell; `tests/assets/probe_empty_search.py`): the answer is empty at once, with no error, even for
//! a sheet that is not there.

use sheetrs_io::Book;
use sheetrs_protocol::{Find, Mode, Query, Scope};
use sheetrs_query::BookQuery;
use std::path::Path;

fn search(extension: &str, find: Find, scope: Vec<Scope>) -> Result<usize, String> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(format!("anchors.{extension}"));
    let query = Query {
        find: Some(find),
        scope: if scope.is_empty() {
            vec![Scope::Workbook]
        } else {
            scope
        },
        ..Query::default()
    };
    Book::from_file(path)
        .unwrap()
        .query(&query)
        .map(|r| r.cells.len())
        .map_err(|e| e.to_string())
}

#[test]
fn the_empty_search_text_finds_nothing_in_every_mode() {
    for extension in ["xlsx", "ods"] {
        // A search that finds cells, to show the file has some.
        assert!(search(extension, Find::literal("a"), vec![]).unwrap() > 0);
        for mode in [Mode::Literal, Mode::Regex, Mode::Wildcards] {
            for (match_case, entire_cells) in [(false, false), (true, false), (false, true)] {
                let find = Find {
                    mode,
                    match_case,
                    entire_cells,
                    ..Find::literal("")
                };
                assert_eq!(search(extension, find.clone(), vec![]), Ok(0), "{mode:?}");
                let nowhere = vec![Scope::Sheet("Nowhere".into())];
                assert_eq!(search(extension, find, nowhere), Ok(0), "{mode:?}");
            }
        }
    }
}
