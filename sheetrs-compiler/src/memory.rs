//! The tools of `sheetc` over bytes and JSON, with no file, process or thread: what a program that
//! has no operating system (WebAssembly in a browser) can run.
//!
//! The command line reads files, a configuration next to the workbook, flags and standard input;
//! these functions take the same things as values. The configuration is [`Options`] (the keys of a
//! `sheetc.toml`, and the declarations of inputs and outputs as the JSON of a `settings_file`), the
//! request is the JSON of the protocol, and the workbook is a [`Book`] (bytes of a given format).
//!
//! | Command | Function |
//! |---|---|
//! | `sheetc calculate -i FILE` (JSON answer) | [`BookCompile::calculate`] with [`Mode::Values`] |
//! | `sheetc calculate -i FILE -o OUT` (the workbook, and the JSON with `--json`) | [`BookCompile::calculate`] with [`Mode::Embedded`] |
//! | `sheetc build FILE --target wasm` | [`BookCompile::build_wasm`] |
//! | `sheetc inspect FILE` | [`BookCompile::inspect`] |
//! | `sheetc verify -i FILE` (`--reference cache`) | [`BookCompile::verify`] |
//! | `sheetc evaluate FORMULA` | [`evaluate`] |

use crate::compile::{Report, compile};
use crate::config::Config;
use crate::error::{Error, Result};
use crate::loader;
use crate::model::Model;
use crate::verify::{self, Verification};
use serde::Serialize;
use serde_json::{Map, Value as Json};
use sheetrs_engine::calculate::{Overrides, ValueRanges};
use sheetrs_engine::ir::{Mode, Program};
use sheetrs_io::Book;

/// What a build is configured with: the keys of a `sheetc.toml` (`fold`, `strict`, `now`, `seed`,
/// `host`, `enable_features`, `locale`, `settings`) and, for a workbook with no settings sheet, the
/// declarations of its inputs and outputs.
#[derive(Debug, Clone, Default)]
pub struct Options {
    /// The configuration, as in a `sheetc.toml`.
    pub config: Config,
    /// The declarations of inputs and outputs as the JSON text of a `settings_file`; they replace
    /// the settings sheet.
    pub declarations: Option<String>,
}

impl Options {
    /// The options of a JSON object: its keys are those of a `sheetc.toml`, plus `declarations`
    /// (the JSON of a `settings_file`, as an object or as text). `null` is the defaults.
    ///
    /// # Errors
    ///
    /// When the value is not an object, a key is not one of the configuration (there is no file
    /// system to read a `settings_file` from: use `declarations`) or a value has the wrong type.
    pub fn from_json(value: &Json) -> Result<Self> {
        let mut map = match value {
            Json::Null => return Ok(Self::default()),
            Json::Object(map) => map.clone(),
            _ => return Err(Error::Config("the options are a JSON object".into())),
        };
        let declarations = match map.remove("declarations") {
            None | Some(Json::Null) => None,
            Some(Json::String(text)) => Some(text),
            Some(other) => Some(other.to_string()),
        };
        if map.contains_key("settings_file") || map.contains_key("settings-file") {
            return Err(Error::Config(
                "`settings_file` names a file, and there is none here: give the declarations as `declarations`"
                    .into(),
            ));
        }
        let config = serde_json::from_value(Json::Object(map))
            .map_err(|e| Error::Config(format!("the options: {e}")))?;
        Ok(Self {
            config,
            declarations,
        })
    }

    fn spec(&self, model: &Model, mode: Mode) -> Result<crate::compile::Spec> {
        let mut spec = match &self.declarations {
            Some(text) => self.config.spec_with_declarations(model, text)?,
            None => self.config.spec(model)?,
        };
        spec.mode = mode;
        Ok(spec)
    }

    /// The values the run overrides: the clock, the seed and the host the options fix.
    #[must_use]
    pub fn overrides(&self) -> Overrides {
        Overrides {
            now: self.config.now,
            rand_seed: self.config.seed,
            host: self.config.host.clone().unwrap_or_default(),
        }
    }
}

/// What a [`Book`] compiles to, runs and says about itself: the tools of `sheetc` as methods of the
/// workbook they read.
pub trait BookCompile {
    /// The program of the workbook, for `mode`, and the report of the build.
    ///
    /// # Errors
    ///
    /// When the workbook cannot be read or compiled.
    fn compile(&self, options: &Options, mode: Mode) -> Result<(Program, Report)>;

    /// The bytes and the extension of the workbook, for the template of a program.
    ///
    /// # Errors
    ///
    /// When a file cannot be read.
    fn template(&self) -> Result<(String, Vec<u8>)>;

    /// The WebAssembly module of the workbook: the program of `mode` (see
    /// [`compile`](Self::compile)) added to the base module of `runtimes` that fits the mode (see
    /// [`crate::wasm_target`]); a workbook mode carries the workbook as its template, and needs the
    /// base that writes workbooks.
    ///
    /// # Errors
    ///
    /// When the workbook cannot be read or compiled, or the module cannot be completed.
    fn build_wasm(
        &self,
        options: &Options,
        mode: Mode,
        runtimes: Runtimes<'_>,
    ) -> Result<(Vec<u8>, Report)>;

    /// Builds the program of the workbook and answers `request` (`{"inputs": {...}, "ranges":
    /// [...]}`), in memory.
    ///
    /// - [`Mode::Values`]: the JSON answer.
    /// - [`Mode::Embedded`]: the workbook with its formulas calculated and the inputs written; the
    ///   JSON answer as well when `with_json` (a request with `ranges` needs it).
    /// - [`Mode::Recalculate`]: every formula calculated again, as it is, the workbook of a
    ///   `sheetrs edit` conversion of formulas to values on stale results; no `sheetc` command
    ///   builds it.
    ///
    /// # Errors
    ///
    /// When the workbook cannot be read or compiled, the request is not valid for the program, the
    /// run fails, or the workbook cannot be written.
    fn calculate(
        &self,
        request: &Json,
        options: &Options,
        mode: Mode,
        with_json: bool,
    ) -> Result<Calculated>;

    /// `sheetc inspect`: the sheets, the cells and formulas, the functions the workbook uses and
    /// which of them the engine lacks.
    ///
    /// ```no_run
    /// use sheetrs_compiler::memory::{BookCompile, Options};
    /// use sheetrs_io::Book;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let book = Book::from_file("model.xlsx")?;
    /// let inspection = book.inspect(&Options::default())?;
    /// println!("{} formulas, unsupported: {:?}", inspection.formulas, inspection.unsupported_functions);
    /// # Ok(()) }
    /// ```
    ///
    /// # Errors
    ///
    /// When the workbook cannot be read or compiled.
    fn inspect(&self, options: &Options) -> Result<Inspection>;

    /// `sheetc verify` with `--reference cache`: the formulas calculated again by the engine against
    /// the values the workbook caches, as JSON (`sheets`, each with its `formulas`, `matched`,
    /// `volatile`, `unparsed` and the `mismatches` with `at`, `formula`, `expected`, `got` and
    /// `kind`).
    ///
    /// # Errors
    ///
    /// When the workbook cannot be read or compiled, or none of the sheets asked for exists.
    fn verify(
        &self,
        options: &Options,
        sheets: &[String],
        tolerance: Option<f64>,
        abs_tolerance: Option<f64>,
    ) -> Result<Json>;
}

impl BookCompile for Book {
    fn compile(&self, options: &Options, mode: Mode) -> Result<(Program, Report)> {
        build(self, options, mode)
    }

    fn template(&self) -> Result<(String, Vec<u8>)> {
        template_of(self)
    }

    fn build_wasm(
        &self,
        options: &Options,
        mode: Mode,
        runtimes: Runtimes<'_>,
    ) -> Result<(Vec<u8>, Report)> {
        build_wasm(self, options, mode, runtimes)
    }

    fn calculate(
        &self,
        request: &Json,
        options: &Options,
        mode: Mode,
        with_json: bool,
    ) -> Result<Calculated> {
        calculate(self, request, options, mode, with_json)
    }

    fn inspect(&self, options: &Options) -> Result<Inspection> {
        inspect(self, options)
    }

    fn verify(
        &self,
        options: &Options,
        sheets: &[String],
        tolerance: Option<f64>,
        abs_tolerance: Option<f64>,
    ) -> Result<Json> {
        verify(self, options, sheets, tolerance, abs_tolerance)
    }
}

/// The program of the workbook, for `mode`, and the report of the build.
///
/// # Errors
///
/// When the workbook cannot be read or compiled.
fn build(source: &Book, options: &Options, mode: Mode) -> Result<(Program, Report)> {
    let model = loader::load(source)?;
    let spec = options.spec(&model, mode)?;
    let (mut program, report) = compile(model, &spec)?;
    program.build.built_at = crate::compile::build_time()?;
    Ok((program, report))
}

/// The two base modules of the WebAssembly target: the runtime with the engine (`values`), and the
/// one that also reads and writes workbooks (`template`), both built by the `sheetrs-rt-wasm` crate.
#[derive(Debug, Clone, Copy)]
pub struct Runtimes<'a> {
    /// The base module that answers JSON (`sheetrs-rt.wasm`).
    pub values: &'a [u8],
    /// The base module that also writes workbooks (`sheetrs-rt-template.wasm`).
    pub template: &'a [u8],
}

/// The bytes and the extension of the workbook, for the template of a program.
///
/// # Errors
///
/// When a file cannot be read.
fn template_of(source: &Book) -> Result<(String, Vec<u8>)> {
    let bytes = match (source.bytes(), source.path()) {
        (Some(data), _) => data.to_vec(),
        (None, Some(path)) => std::fs::read(path).map_err(|e| Error::Io {
            path: path.to_path_buf(),
            source: e,
        })?,
        (None, None) => unreachable!("a book is a file or bytes"),
    };
    Ok((source.format().extension().to_owned(), bytes))
}

/// The WebAssembly module of the workbook: the program of `mode` (see [`BookCompile::compile`]) added to the base
/// module of `runtimes` that fits the mode (see [`crate::wasm_target`]); a workbook mode carries the
/// workbook as its template, and needs the base that writes workbooks.
///
/// # Errors
///
/// When the workbook cannot be read or compiled, or the module cannot be completed.
fn build_wasm(
    source: &Book,
    options: &Options,
    mode: Mode,
    runtimes: Runtimes<'_>,
) -> Result<(Vec<u8>, Report)> {
    let (mut program, report) = build(source, options, mode)?;
    let template = mode
        .is_workbook()
        .then(|| template_of(source))
        .transpose()?;
    let payload = sheetrs_engine::bundle::payload(
        &mut program,
        template.as_ref().map(|(e, b)| (e.as_str(), b.as_slice())),
    )?;
    let base = if mode.is_workbook() {
        runtimes.template
    } else {
        runtimes.values
    };
    Ok((crate::wasm_target::complete(base, &payload)?, report))
}

/// What [`BookCompile::calculate`] gives back.
#[derive(Debug, Default)]
pub struct Calculated {
    /// The answer of the protocol: the `ranges` the request asked for, or the declared outputs.
    /// `None` when the mode is a workbook one and the answer was not asked for.
    pub answer: Option<ValueRanges>,
    /// The workbook with the answer written into it, in a workbook mode.
    pub workbook: Option<Vec<u8>>,
    /// The cells written into the workbook.
    pub written: usize,
    /// What the compiler noted (locale, macros) and the diagnostics of the program.
    pub diagnostics: Vec<String>,
}

/// Builds the program of `source` and answers `request` (`{"inputs": {...}, "ranges": [...]}`),
/// in memory.
///
/// - [`Mode::Values`]: the JSON answer.
/// - [`Mode::Embedded`]: the workbook with its formulas calculated and the inputs written; the JSON
///   answer as well when `with_json` (a request with `ranges` needs it).
/// - [`Mode::Recalculate`]: every formula calculated again, as it is, the workbook of a `sheetrs edit` conversion of formulas to values on stale results; no `sheetc` command builds it.
///
/// # Errors
///
/// When the workbook cannot be read or compiled, the request is not valid for the program, the run
/// fails, or the workbook cannot be written.
fn calculate(
    source: &Book,
    request: &Json,
    options: &Options,
    mode: Mode,
    with_json: bool,
) -> Result<Calculated> {
    let (program, _) = build(source, options, mode)?;
    let overrides = options.overrides();
    let diagnostics = program.diagnostics.clone();
    if !mode.is_workbook() {
        let answer = program.handle(request, &overrides)?;
        return Ok(Calculated {
            answer: Some(answer),
            diagnostics,
            ..Calculated::default()
        });
    }
    if mode == Mode::Recalculate && with_json {
        return Err(Error::Config(
            "a workbook recalculated as it is gives no JSON".into(),
        ));
    }
    if with_json {
        let mut workbook = Vec::new();
        let (report, answer) = sheetrs_answer::fill_template_to_with_answer(
            source,
            &program,
            None,
            request,
            &overrides,
            &mut workbook,
        )?;
        Ok(Calculated {
            answer: Some(answer),
            workbook: Some(workbook),
            written: report.cells,
            diagnostics,
        })
    } else {
        let mut workbook = Vec::new();
        let report = sheetrs_answer::fill_template_to(
            source,
            &program,
            None,
            request,
            &overrides,
            &mut workbook,
        )?;
        Ok(Calculated {
            answer: None,
            workbook: Some(workbook),
            written: report.cells,
            diagnostics,
        })
    }
}

/// What `sheetc inspect` says about a workbook.
#[derive(Debug, Serialize)]
pub struct Inspection {
    /// The names of the sheets.
    pub sheets: Vec<String>,
    /// The null date of the workbook (`1899-12-30`, `1900-01-01` or `1904-01-01`).
    pub date_system: String,
    /// Cells with a value or a formula.
    pub cells: usize,
    /// Cells with a formula.
    pub formulas: usize,
    /// Groups of cells that read one another.
    pub circular_groups: usize,
    /// `INDIRECT` or `OFFSET` are used: the whole workbook is reachable.
    pub run_time_references: bool,
    /// Functions the workbook uses that the engine does not have (and `LibreOffice` does not know
    /// either).
    pub unsupported_functions: Vec<String>,
    /// The functions used, by number of calls, most used first.
    pub functions: Vec<FunctionUse>,
    /// What the compiler says about the workbook while it reads it.
    pub diagnostics: Vec<String>,
}

/// A function a workbook uses.
#[derive(Debug, Serialize)]
pub struct FunctionUse {
    /// Its name.
    pub name: String,
    /// The number of calls.
    pub calls: usize,
    /// Whether the engine has it.
    pub implemented: bool,
}

/// `sheetc inspect`: the sheets, the cells and formulas, the functions the workbook uses and which
/// of them the engine lacks.
///
/// # Errors
///
/// When the workbook cannot be read or compiled.
fn inspect(source: &Book, options: &Options) -> Result<Inspection> {
    let model = loader::load(source)?;
    let mut config = options.config.clone();
    config.strict = Some(false);
    let mut spec = match &options.declarations {
        Some(text) => config.spec_with_declarations(&model, text)?,
        None => config.spec(&model)?,
    };
    spec.fold = false;
    spec.inspect = true;
    let (program, report) = compile(&model, &spec)?;
    let mut uses: Vec<(&String, &usize)> = report.function_uses.iter().collect();
    uses.sort_by(|a, b| b.1.cmp(a.1).then(a.0.cmp(b.0)));
    Ok(Inspection {
        sheets: model.sheets.clone(),
        date_system: model.date_system.to_string(),
        cells: report.cells_total,
        formulas: report.formulas_total,
        circular_groups: report.cycles,
        run_time_references: report.dynamic_refs,
        unsupported_functions: report.unsupported.iter().cloned().collect(),
        functions: uses
            .into_iter()
            .map(|(name, calls)| FunctionUse {
                name: name.clone(),
                calls: *calls,
                implemented: !report.unsupported.contains(name),
            })
            .collect(),
        diagnostics: program.diagnostics.iter().take(20).cloned().collect(),
    })
}

/// `sheetc verify` with `--reference cache`: the formulas calculated again by the engine against the
/// values the workbook caches, as JSON (`sheets`, each with its `formulas`, `matched`, `volatile`,
/// `unparsed` and the `mismatches` with `at`, `formula`, `expected`, `got` and `kind`).
///
/// # Errors
///
/// When the workbook cannot be read or compiled, or none of the sheets asked for exists.
fn verify(
    source: &Book,
    options: &Options,
    sheets: &[String],
    tolerance: Option<f64>,
    abs_tolerance: Option<f64>,
) -> Result<Json> {
    let model = loader::load(source)?;
    let mut config = options.config.clone();
    config.strict = Some(false);
    let spec = match &options.declarations {
        Some(text) => config.spec_with_declarations(&model, text)?,
        None => config.spec(&model)?,
    };
    let defaults = verify::Options::default();
    let opts = verify::Options {
        sheets: sheets.to_vec(),
        tolerance: tolerance.unwrap_or(defaults.tolerance),
        abs_tolerance: abs_tolerance.unwrap_or(defaults.abs_tolerance),
    };
    let result = verify::verify(&model, &spec, &opts, None)?;
    Ok(verification_json(&result))
}

fn verification_json(v: &Verification) -> Json {
    let sheets: Vec<Json> = v
        .sheets
        .iter()
        .map(|s| {
            let mismatches: Vec<Json> = s
                .mismatches
                .iter()
                .map(|m| {
                    let mut o = Map::new();
                    o.insert("at".into(), Json::from(m.at.clone()));
                    o.insert("formula".into(), Json::from(m.formula.clone()));
                    o.insert("expected".into(), Json::from(m.expected.clone()));
                    o.insert("got".into(), Json::from(m.got.clone()));
                    o.insert("kind".into(), Json::from(m.kind));
                    Json::Object(o)
                })
                .collect();
            let mut o = Map::new();
            o.insert("name".into(), Json::from(s.name.clone()));
            o.insert("formulas".into(), Json::from(s.formulas));
            o.insert("matched".into(), Json::from(s.matched));
            o.insert("volatile".into(), Json::from(s.volatile));
            o.insert("unparsed".into(), Json::from(s.unparsed));
            o.insert("mismatches".into(), Json::Array(mismatches));
            Json::Object(o)
        })
        .collect();
    let mut o = Map::new();
    o.insert("mismatches".into(), Json::from(v.mismatches()));
    o.insert("sheets".into(), Json::Array(sheets));
    Json::Object(o)
}

/// `sheetc evaluate`: the value of one formula, with the regional settings of `locale` (a BCP-47
/// tag, `en` by default).
///
/// # Errors
///
/// When the locale is not valid or the formula cannot be compiled.
pub fn evaluate(formula: &str, locale: Option<&str>, options: &Options) -> Result<Json> {
    use crate::compile::{BindingSpec, Spec};
    use crate::locale::{LocaleOverride, from_tag};
    let mut model = Model::new();
    let sheet = model.add_sheet("Sheet1");
    model.set_formula(sheet, "A1", formula);
    let spec = Spec {
        outputs: vec![BindingSpec::parse("Sheet1!A1")],
        locale: Some(from_tag(locale.unwrap_or("en"), &LocaleOverride::default())?.0),
        now: options.config.now,
        rand_seed: options.config.seed,
        host: options.config.host.clone().unwrap_or_default(),
        ..Spec::default()
    };
    let (program, _) = compile(model, &spec)?;
    let out = program.calculate(&Json::Null, &Overrides::default())?;
    Ok(out["Sheet1!A1"].clone())
}
