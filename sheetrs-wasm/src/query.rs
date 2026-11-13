//! The reads of a `Book`: `query`, `range` and `find`.

use crate::book::{Book, failure};
use crate::js::{Targets, from_js, from_js_object, targets, to_js};
use serde::Deserialize;
use serde_json::{Value as Json, json};
use sheetrs_protocol::{Find, Mode, Property, Query, Scope, SearchIn};
use sheetrs_query::{Block, BookQuery, FindOptions, Major, RangeOptions};
use wasm_bindgen::prelude::*;

#[derive(Debug, Deserialize, Clone, Copy, Default)]
#[serde(rename_all = "lowercase")]
enum MajorArg {
    #[default]
    Rows,
    Columns,
}

impl From<MajorArg> for Major {
    fn from(major: MajorArg) -> Self {
        match major {
            MajorArg::Rows => Self::Rows,
            MajorArg::Columns => Self::Columns,
        }
    }
}

/// The options of `range`: `properties` (`"value"`, `"formula"`, `"number_format"`, `"display"`,
/// `"comment"`; the value by default), `major` (`"rows"` or `"columns"`) and `dates` (the `dates` of
/// the `Book` by default).
#[derive(Debug, Default, Deserialize)]
#[serde(default, rename_all = "camelCase", deny_unknown_fields)]
struct RangeArgs {
    properties: Option<Vec<Property>>,
    major: MajorArg,
    dates: Option<bool>,
}

/// The options of `find`: the search (`mode`, `matchCase`, `entireCells`, `searchIn`,
/// `includeFormulas`, `formatted`, `filtered`) and what it returns (`scope`, `properties`, `by`,
/// `dates`).
#[derive(Debug, Default, Deserialize)]
#[serde(default, rename_all = "camelCase", deny_unknown_fields)]
#[allow(clippy::struct_excessive_bools)] // the options of a search, as the protocol has them
struct FindArgs {
    mode: Mode,
    match_case: bool,
    entire_cells: bool,
    search_in: SearchIn,
    include_formulas: bool,
    formatted: bool,
    filtered: bool,
    scope: Option<Targets>,
    properties: Option<Vec<Property>>,
    by: MajorArg,
    dates: Option<bool>,
}

fn block_json(block: &Block) -> Json {
    let mut object = serde_json::Map::new();
    object.insert("range".into(), json!(block.range));
    let mut put = |key: &str, grid: Option<Json>| {
        if let Some(grid) = grid {
            object.insert(key.into(), grid);
        }
    };
    put("values", block.values.as_ref().map(|g| json!(g)));
    put("formulas", block.formulas.as_ref().map(|g| json!(g)));
    put(
        "numberFormats",
        block.number_formats.as_ref().map(|g| json!(g)),
    );
    put("displays", block.displays.as_ref().map(|g| json!(g)));
    put("comments", block.comments.as_ref().map(|g| json!(g)));
    Json::Object(object)
}

#[wasm_bindgen]
impl Book {
    /// Reads the cells a query asks for. `query` is the object of the protocol (the JSON that
    /// `sheetrs query` builds from its flags): `scope`, `properties`, `find`, `locale` and `dates`
    /// (the `dates` of the `Book` when absent). Returns `{cells: [...]}`, as `sheetrs query` prints.
    ///
    /// # Errors
    ///
    /// A message when operations are pending, the workbook cannot be read, the query is not valid or
    /// the search is one the matcher does not reproduce.
    pub fn query(&self, query: JsValue) -> Result<JsValue, JsValue> {
        self.queryable()?;
        let mut query: Query = if query.is_undefined() || query.is_null() {
            Query::default()
        } else {
            from_js(query, "The query")?
        };
        query.dates |= self.dates;
        let response = self
            .inner
            .query(&query)
            .map_err(|e| failure("Query error", e))?;
        to_js(&response)
    }

    /// The cells of one or several ranges (a text or a list of them: `Data!A1:C10`, `Data!A:C`, a
    /// defined name), read in one pass: one block per range, in the order asked
    /// (`{range, values, formulas, numberFormats, displays, comments}`, the grids that were asked
    /// for, as lists of rows or of columns). `options`: `properties`, `major` (`"rows"`,
    /// `"columns"`) and `dates`.
    ///
    /// # Errors
    ///
    /// A message when operations are pending, a range is not valid or the workbook cannot be read.
    pub fn range(&self, ranges: JsValue, options: JsValue) -> Result<JsValue, JsValue> {
        self.queryable()?;
        let ranges = targets(ranges, "The ranges")?;
        let args: RangeArgs = if options.is_undefined() || options.is_null() {
            RangeArgs::default()
        } else {
            from_js_object(options, "The options", &["properties", "major", "dates"])?
        };
        let defaults = RangeOptions::default();
        let options = RangeOptions {
            properties: args.properties.unwrap_or(defaults.properties),
            major: args.major.into(),
            dates: args.dates.unwrap_or(self.dates),
        };
        let blocks = self
            .inner
            .range(&ranges, &options)
            .map_err(|e| failure("Query error", e))?;
        to_js(&blocks.iter().map(block_json).collect::<Vec<_>>())
    }

    /// The cells that match a search, as Find & Replace of `LibreOffice` does: a list of
    /// `{sheet, cell, row, col, value, ...}` with the properties asked for. `options` are those of
    /// the search (`mode`, `matchCase`, `entireCells`, `searchIn`, `includeFormulas`, `formatted`,
    /// `filtered`) and of the answer (`scope`, `properties`, `by`, `dates`).
    ///
    /// # Errors
    ///
    /// A message when operations are pending, the workbook cannot be read or the search is one the
    /// matcher does not reproduce.
    pub fn find(&self, text: &str, options: JsValue) -> Result<JsValue, JsValue> {
        self.queryable()?;
        let args: FindArgs = if options.is_undefined() || options.is_null() {
            FindArgs::default()
        } else {
            from_js_object(
                options,
                "The options",
                &[
                    "mode",
                    "matchCase",
                    "entireCells",
                    "searchIn",
                    "includeFormulas",
                    "formatted",
                    "filtered",
                    "scope",
                    "properties",
                    "by",
                    "dates",
                ],
            )?
        };
        let find = Find {
            text: text.to_string(),
            mode: args.mode,
            match_case: args.match_case,
            entire_cells: args.entire_cells,
            search_in: args.search_in,
            include_formulas: args.include_formulas,
            formatted: args.formatted,
            filtered: args.filtered,
        };
        let defaults = FindOptions::default();
        let options = FindOptions {
            scope: args
                .scope
                .map(Targets::into_vec)
                .unwrap_or_default()
                .iter()
                .map(|s| Scope::from(s.as_str()))
                .collect(),
            properties: args.properties.unwrap_or(defaults.properties),
            by: args.by.into(),
            dates: args.dates.unwrap_or(self.dates),
        };
        let cells = self
            .inner
            .find(&find, &options)
            .map_err(|e| failure("Query error", e))?;
        to_js(&cells)
    }
}
