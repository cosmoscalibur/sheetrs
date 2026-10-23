//! Every formula in the repository's real workbooks must parse (but for the
//! documented unsupported syntax) and print to a fixed point.

use sheetrs::reader::read_workbook;
use sheetrs_formula::{Expr, parse, parse_spanned};
use std::path::Path;

fn files() -> Vec<std::path::PathBuf> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("..");
    let mut out = vec![
        root.join("tests/minimal_test.xlsx"),
        root.join("tests/minimal_test.ods"),
    ];
    if let Ok(dir) = std::fs::read_dir(root.join("sheetrs-compiler/tests/assets")) {
        out.extend(
            dir.flatten()
                .map(|e| e.path())
                .filter(|p| matches!(p.extension().and_then(|e| e.to_str()), Some("xlsx" | "ods"))),
        );
    }
    out
}

#[test]
fn workbook_formulas_roundtrip() {
    let (mut total, mut unsupported) = (0usize, Vec::new());
    for path in files() {
        let wb = read_workbook(&path).unwrap();
        for sheet in &wb.sheets {
            for cell in sheet.cells.values() {
                let Some(f) = cell.formula.as_deref() else {
                    continue;
                };
                total += 1;
                match parse_spanned(f) {
                    Ok(formula) => {
                        let e = &formula.expr;
                        let printed = e.to_string();
                        assert_eq!(&parse(&printed).unwrap(), e, "{}: {f}", path.display());
                        // The span of every reference is its exact source text.
                        let refs = e.references();
                        assert_eq!(refs.len(), formula.ref_spans.len(), "{f}");
                        for (r, span) in refs.iter().zip(&formula.ref_spans) {
                            let text = &formula.source[span.clone()];
                            assert_eq!(
                                parse(text).unwrap(),
                                Expr::Ref((*r).clone()),
                                "{}: `{text}` in `{f}`",
                                path.display()
                            );
                        }
                    }
                    // External-workbook references (`[1]Sheet1!B3`) are not
                    // part of the grammar.
                    Err(_) if f.contains('[') => unsupported.push(f.to_string()),
                    Err(err) => panic!("{}: `{f}`: {err}", path.display()),
                }
            }
        }
    }
    assert!(
        total > 100,
        "expected the fixtures to hold many formulas, got {total}"
    );
    assert!(unsupported.len() < 10, "{unsupported:?}");
}
