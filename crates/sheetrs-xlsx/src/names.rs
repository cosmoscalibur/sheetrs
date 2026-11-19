//! Defined names of the workbook part.

use sheetrs_core::{NamedExpression, Workbook};

use crate::Error;
use crate::lexical;
use crate::parts::{DefinedName, WORKBOOK_PART};

/// The prefix of the names that Excel defines for its own uses, such as the print area.
const BUILT_IN_PREFIX: &str = "_xlnm.";

/// The names Excel defines for its own uses, written after `_xlnm.`.
pub(crate) const BUILT_IN_NAMES: &[&str] = &[
    "Consolidate_Area",
    "Auto_Open",
    "Auto_Close",
    "Extract",
    "Database",
    "Criteria",
    "Print_Area",
    "Print_Titles",
    "Recorder",
    "Data_Form",
    "Auto_Activate",
    "Auto_Deactivate",
    "Sheet_Title",
];

/// The built-in name of the range an automatic filter covers, which `LibreOffice` keeps as a
/// property of the sheet and not as a name.
const FILTER_DATABASE: &str = "_FilterDatabase";

/// Adds the defined names to `workbook`, whose sheets are already loaded, as `LibreOffice` loads
/// them, and returns the elements it does not keep as names, as written. A name that starts with
/// `_xlnm.` starts with `Excel_BuiltIn_` instead, except the one of an automatic filter, which is
/// not kept as a name; a name whose expression is a reference into another document is dropped.
///
/// # Errors
///
/// [`Error::Malformed`] when a name belongs to a sheet that is not in the workbook, which
/// `LibreOffice` does not load.
pub(crate) fn load_names(
    workbook: &mut Workbook,
    defined: &[DefinedName],
    external_documents: usize,
) -> Result<Vec<String>, Error> {
    let sheets = workbook.sheets().len();
    let mut unread = Vec::new();
    for entry in defined {
        let Some(name) = entry.name.as_deref().filter(|name| !name.is_empty()) else {
            continue;
        };
        let scope = match entry.local_sheet.as_deref().map(sheetrs_io::integer_prefix) {
            Some(index) if index >= 0 => {
                let index = usize::try_from(index).unwrap_or(usize::MAX);
                if index >= sheets {
                    return Err(Error::malformed(
                        WORKBOOK_PART,
                        format!("the name {name:?} belongs to a sheet that is not there"),
                    ));
                }
                Some(index)
            }
            _ => None,
        };
        let built_in = name
            .strip_prefix(BUILT_IN_PREFIX)
            .filter(|rest| *rest == FILTER_DATABASE || BUILT_IN_NAMES.contains(rest));
        if built_in == Some(FILTER_DATABASE) {
            unread.push(format!(
                "<definedName {}>{}</definedName>",
                entry.attributes,
                escape_text(&entry.text)
            ));
            continue;
        }
        let name = built_in.map_or_else(|| name.to_owned(), |rest| format!("Excel_BuiltIn_{rest}"));
        let expression = entry.text.trim();
        let expression = expression.strip_prefix('=').unwrap_or(expression).trim();
        if lexical::is_external_reference(expression) {
            continue;
        }
        for index in lexical::external_indices(expression) {
            // `LibreOffice` names a document that the package does not list by the number.
            if index == 0 || index > external_documents {
                workbook.add_external_link(&index.to_string());
            }
        }
        workbook.define_name(NamedExpression {
            name,
            scope,
            expression: expression.to_owned(),
            hidden: entry.hidden.as_deref().is_some_and(lexical::workbook_flag),
        });
    }
    Ok(unread)
}

fn escape_text(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn entry(name: &str, local: Option<&str>, text: &str) -> DefinedName {
        DefinedName {
            name: Some(name.to_owned()),
            local_sheet: local.map(str::to_owned),
            hidden: None,
            text: text.to_owned(),
            attributes: format!("name=\"{name}\""),
        }
    }

    fn workbook() -> Workbook {
        let mut workbook = Workbook::new();
        workbook.sheet_builder("S1").finish();
        workbook.sheet_builder("S2").finish();
        workbook
    }

    #[test]
    fn built_in_names_change_their_prefix_except_the_filter_range() -> Result<(), Error> {
        let mut book = workbook();
        let unread = load_names(
            &mut book,
            &[
                entry("_xlnm.Print_Area", Some("0"), "S1!$A$1"),
                entry("_xlnm.Print_Titles", None, "S1!$1:$1"),
                entry("_xlnm._FilterDatabase", Some("0"), "S1!$A$1:$B$2"),
            ],
            0,
        )?;
        assert_eq!(unread.len(), 1);
        let names: Vec<_> = book.names().iter().map(|name| name.name.as_str()).collect();
        assert_eq!(
            names,
            ["Excel_BuiltIn_Print_Area", "Excel_BuiltIn_Print_Titles"]
        );
        Ok(())
    }

    #[test]
    fn the_sheet_of_a_name_is_read_as_an_integer_prefix() -> Result<(), Error> {
        let mut book = workbook();
        load_names(
            &mut book,
            &[
                entry("A", Some("x"), "1"),
                entry("B", Some("1x"), "1"),
                entry("C", Some("-1"), "1"),
            ],
            0,
        )?;
        let scopes: Vec<_> = book.names().iter().map(|name| name.scope).collect();
        assert_eq!(scopes, vec![Some(0), Some(1), None]);
        Ok(())
    }

    #[test]
    fn a_name_of_a_missing_sheet_is_refused() {
        let mut book = workbook();
        assert!(load_names(&mut book, &[entry("A", Some("9"), "1")], 0).is_err());
    }

    #[test]
    fn expressions_lose_their_equal_sign_and_spaces() -> Result<(), Error> {
        let mut book = workbook();
        load_names(&mut book, &[entry("A", None, "  =S1!$A$1 ")], 0)?;
        assert_eq!(book.names()[0].expression, "S1!$A$1");
        Ok(())
    }
}
