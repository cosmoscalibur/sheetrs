//! `sheetc`: build, calculate and inspect spreadsheet programs.

use sheetrs_compiler::config::{self, Config};
use sheetrs_compiler::locale::LocaleOverride;
use sheetrs_compiler::{Model, Spec, compile, loader};
use sheetrs_engine::Program;
use sheetrs_engine::bundle;
use sheetrs_engine::cli::{self, Args, parse_overrides};
use sheetrs_io::Book;
use sheetrs_protocol::{Diagnostic, LogFormat, Logger};
use std::path::{Path, PathBuf};
use std::sync::OnceLock;

const HELP: &str = "\
sheetc — compile XLSX/ODS spreadsheets into executables

USAGE
  sheetc build -i FILE [-o OUT] [options]
      Compile FILE into a standalone executable (`--target wasm`: a WebAssembly module).
      The workbook declares its inputs, outputs and metadata in a settings sheet (see below).
      `--emit-ir PATH` also writes the program as JSON.
  sheetc calculate -i FILE|PROGRAM.json [--request REQUEST] [--range RANGE]... [-o OUT | --in-place]
                   [--json] [options]
      Compile in memory and calculate immediately. The inputs and the outputs of the cells are
      declared (settings sheet or file) or deduced. `--request` is the request (`inputs`, `ranges`,
      `majorDimension`): a JSON object as text, `@FILE`, or `-` for the standard input (read only
      then); `--range` adds a range to `ranges` (repeatable). Without `-o` or `--in-place` the answer is JSON
      on the standard output. With `-o OUT` the answer is a copy of FILE with the inputs given and
      the results written into it; `--in-place` replaces FILE by that workbook (written beside it and
      renamed, so a failure leaves FILE as it was). `-o` never overwrites FILE. `--json` adds the
      JSON to the workbook, from the same calculation. With nothing declared or given, every formula
      is calculated again and the workbook is written as it is.
  sheetc evaluate \"=FORMULA\" [--locale TAG] [--now N] [--seed N]
      Evaluate one formula.
  sheetc --version
      The version of the compiler.
  sheetc inspect -i FILE [--top N]
      Show what the compiler sees: sizes, unsupported functions, cycles, and the
      functions the formulas use, most used first, with whether the engine has them.
  sheetc verify -i FILE [--sheet NAME]... [--tolerance X] [--abs-tolerance X] [--max N]
                [--reference cache|libreoffice] [build options]
      Recompute the formulas of FILE (all sheets, or the named ones) with the
      values the file holds and compare them with a reference. Reports what differs; exit code 1
      if anything does. The reference is the values the file cached (`--reference cache`, the
      default) or a full recalculation by LibreOffice ($SOFFICE or `soffice`) under the build
      locale (`--reference libreoffice`).

The workbook is `-i`, `--input` or `--file`; the file that is written is `-o` or `--output`.

LOGS (every command of `sheetc`; not the executables, which answer in JSON)
  Warnings and errors are written on the standard error as text. `-q` keeps only the errors; `-v` adds
  what the command did (what it wrote, the summary of a build), `-vv` the detail. `--log-format json`
  writes each as a JSON record per line: {\"level\", \"code\", \"message\", \"context\"}.

SETTINGS SHEET
  First row: key, value, description, meta_type. A row with meta_type `input`
  or `output` declares the range in `key` (`Sheet1!B2:C9`); any other type is
  free metadata, and a row without type is the category `extra`. The executable
  lists them with `--meta`, after the properties of the document. Without the sheet
  the outputs are the formula cells and the inputs are deduced.

BUILD OPTIONS (each also a key of sheetc.toml; the command line wins)
  --config PATH        configuration file (default: sheetc.toml next to FILE)
  --settings SHEET     settings sheet (default: `settings`, optional)
  --settings-file PATH JSON with the same declarations (inputs, outputs, meta), for
                       a workbook without a settings sheet: it compiles and
                       calculates as it is, inputs included. Replaces the sheet
  --locale TAG         regional settings, any BCP-47 tag (default `en`); fixed
                       at build time. More in the [locale] table of the file.
  --enable-features L  compatibility features on, comma separated (all are
                       off by default). LONGTEXT joins
                       the pieces of `_xlfn._LONGTEXT(a,b,...)` (a text literal
                       over 255 characters written by LibreOffice 26.2);
                       constant text pieces only
  --no-fold            keep every reachable formula (no build-time evaluation)
  --no-strict          compile a workbook with macros (reported as a warning);
                       by default (strict) it is an error. --strict undoes a
                       `strict = false` of the file
  --now SERIAL   --seed N
  --system S  --osversion S  --release S  --directory PATH  --filename NAME
                       fix what INFO and CELL(filename) say (else the machine
                       and the executable's directory, at run time); [host] in
                       the configuration file
  --runtime PATH       runtime stub to embed into (default: sheetrs-rt next to
                       this binary, else this binary; with `--target wasm`, the base module
                       (`sheetrs-rt.wasm`, or `sheetrs-rt-template.wasm` for `--embed-template`)
                       next to this binary: built by scripts/build-wasm-rt.sh, never replaced)
  --target native|wasm `wasm` writes a WebAssembly module (OUT.wasm) instead of an executable: the
                       base module plus the program, to be instantiated by a browser or Node
                       with no glue (see sheetrs-wasm/README.md for its exports)
  --embed-template     (only in `build`) the executable carries the workbook, which it does not have
                       otherwise, so that it can write a workbook as its answer: `-o OUT` writes a
                       copy of the workbook with the inputs given (`--request`) and the results in its
                       cells; without `-o` the answer is JSON, as in any executable; `-o OUT --json`
                       gives both from one calculation. Inputs and outputs are declared or deduced as
                       in any executable; the formulas outside the program keep the result the
                       template stored. The executable is made from `sheetc` itself, which writes
                       workbooks.
  The date system, iteration and search settings come from the document and
  are not options. `evaluate` has no document and accepts
  --date-system 1899|1900|1904|YYYY-MM-DD and --iterate[=N[,EPS]].

Executables accept: --request, --range RANGE, --serve, --pretty, --meta, --info, --now, --seed,
--system, --osversion, --release, --directory, --filename. One built with --embed-template also
accepts -o/--output OUT and --json (the workbook, and the JSON with it). An executable has no file
to replace: it has no --in-place.";

fn main() {
    let args = Args(std::env::args().skip(1).collect());
    #[cfg(feature = "profile")]
    let profiler = std::env::var("SHEETC_FLAMEGRAPH").ok().and_then(|path| {
        pprof::ProfilerGuardBuilder::default()
            .frequency(997)
            .blocklist(&["libc", "libgcc", "pthread", "vdso"])
            .build()
            .ok()
            .map(|guard| (path, guard))
    });
    let code = cli::with_big_stack(move || real_main(args));
    #[cfg(feature = "profile")]
    if let Some((path, guard)) = profiler
        && let Ok(report) = guard.report().build()
        && let Ok(file) = std::fs::File::create(&path)
    {
        let _ = report.flamegraph(file);
        eprintln!("flamegraph written to {path}");
    }
    std::process::exit(code);
}

/// Where the diagnostics of the command go: set once, from the flags of the command line.
static LOGGER: OnceLock<Logger> = OnceLock::new();

fn log(diagnostic: &Diagnostic) {
    LOGGER.get_or_init(Logger::default).emit(diagnostic);
}

/// The logger the flags `-q`, `-v` and `--log-format` ask for; they are taken out of the arguments.
fn logger_of(args: &mut Args) -> Result<Logger, Fail> {
    let quiet = args.flag("-q") || args.flag("--quiet");
    let mut verbose = 0u8;
    while args.flag("-v") || args.flag("--verbose") {
        verbose = verbose.saturating_add(1);
    }
    while args.flag("-vv") {
        verbose = verbose.saturating_add(2);
    }
    let format = match args.value("--log-format")? {
        Some(name) => name.parse::<LogFormat>().map_err(Fail::Usage)?,
        None => LogFormat::Human,
    };
    if quiet && verbose > 0 {
        return Err(Fail::Usage("`-q` and `-v` cannot be used together".into()));
    }
    Ok(Logger::new(quiet, verbose, format))
}

fn real_main(mut args: Args) -> i32 {
    if let Some(e) = bundle::read_embedded() {
        return match embedded_workbook(&e, args) {
            Ok(code) => code,
            Err(Fail::Usage(m)) => cli::usage_error(&m),
            Err(Fail::Run(m)) => {
                eprintln!("error: {m}");
                1
            }
        };
    }
    if args.0.is_empty() || args.flag("--help") || args.flag("-h") {
        println!("{HELP}");
        return 0;
    }
    if args.flag("--version") || args.flag("-V") {
        println!("sheetc {}", env!("CARGO_PKG_VERSION"));
        return 0;
    }
    let cmd = args.0.remove(0);
    match logger_of(&mut args) {
        Ok(logger) => {
            let _ = LOGGER.set(logger);
        }
        Err(Fail::Usage(m)) => return cli::usage_error(&m),
        Err(Fail::Run(m)) => {
            log(&Diagnostic::error("COMMAND_FAILED", m));
            return 1;
        }
    }
    let result = match cmd.as_str() {
        "build" => build(args),
        "calculate" => calculate(args),
        "evaluate" => evaluate(args),
        "inspect" => inspect(args),
        "verify" => verify(args),
        other => Err(Fail::Usage(format!("unknown command `{other}`\n\n{HELP}"))),
    };
    match result {
        Ok(code) => code,
        Err(Fail::Usage(m)) => cli::usage_error(&m),
        Err(Fail::Run(m)) => {
            log(&Diagnostic::error("COMMAND_FAILED", m));
            1
        }
    }
}

enum Fail {
    Usage(String),
    Run(String),
}

impl From<String> for Fail {
    fn from(s: String) -> Self {
        Self::Usage(s)
    }
}

impl From<sheetrs_compiler::Error> for Fail {
    fn from(e: sheetrs_compiler::Error) -> Self {
        Self::Run(e.to_string())
    }
}

/// A wrong command line is a usage error (exit 2); the rest failed while running.
impl From<sheetrs_engine::Error> for Fail {
    fn from(e: sheetrs_engine::Error) -> Self {
        match e {
            sheetrs_engine::Error::Usage(message) => Self::Usage(message),
            other => Self::Run(other.to_string()),
        }
    }
}

/// An executable made from a workbook: the answer as JSON, or, when it carries its workbook as the
/// template, the workbook with the answer written into it (`-o OUT`), the JSON, or both. It has no
/// file to replace, so it has no `--in-place`.
fn embedded_workbook(e: &bundle::Embedded, mut args: Args) -> Result<i32, Fail> {
    let Some(template) = &e.template else {
        if output_file(&mut args)?.is_some() {
            return Err(Fail::Usage(
                "this executable carries no workbook: build it with --embed-template to write one"
                    .into(),
            ));
        }
        return Ok(cli::calculate_program(&e.program, e.consts.as_ref(), args));
    };
    let source = Book::from_bytes(template.bytes.clone())
        .map_err(|e| Fail::Run(format!("the template: {e}")))?;
    let out = output_file(&mut args)?;
    let json = args.flag("--json");
    embedded_answer(
        &e.program,
        e.consts.as_ref(),
        &source,
        out,
        None,
        json,
        args,
    )
}

/// The answer of a program that carries its workbook as the template, to the runtime options in
/// `args` (those of any executable, plus `-o OUT` and `--json`, already taken):
///
/// - without `out`, the JSON answer on stdout, as in any executable (`--range`, `--pretty`,
///   `--serve`, `--meta`, `--info`);
/// - with `out`, the workbook written to it and nothing on stdout;
/// - with `out` and `--json`, both, from one calculation (`--serve` is the JSON of many requests and
///   does not write a workbook).
///
/// `shown` is what the message `wrote …` names when it is not `out` (the file a `--in-place` replaces).
fn embedded_answer(
    program: &Program,
    consts: Option<&sheetrs_engine::consts::StoredConsts>,
    template: &Book,
    out: Option<String>,
    shown: Option<&str>,
    json: bool,
    mut args: Args,
) -> Result<i32, Fail> {
    // The metadata is answered as in any executable.
    if args
        .0
        .iter()
        .any(|a| matches!(a.as_str(), "--meta" | "--info" | "--version"))
    {
        return Ok(cli::calculate_program(program, consts, args));
    }
    let base = consts.map(|c| c as &dyn sheetrs_engine::engine::ConstSource);
    let Some(out) = out else {
        // `--json` alone is what it already does.
        return Ok(cli::calculate_program(program, consts, args));
    };
    if args.flag("--serve") {
        return Err(Fail::Usage(
            "--serve answers JSON, one request per line: it does not write a workbook".into(),
        ));
    }
    let pretty = args.flag("--pretty");
    let ranges = args.values("--range")?;
    if !json && (pretty || !ranges.is_empty()) {
        return Err(Fail::Usage(
            "--pretty and --range are for the JSON answer: add --json to have it as well as the workbook"
                .into(),
        ));
    }
    let input = args.value("--request")?;
    let o = parse_overrides(&mut args)?;
    if let Some(extra) = args.0.first() {
        return Err(Fail::Usage(format!(
            "unexpected argument `{extra}`: an executable that carries its workbook answers with it"
        )));
    }
    let named = shown.unwrap_or(&out).to_string();
    if !json {
        let request = cli::read_request(input.as_deref())?;
        return write_workbook(program, base, template, &request, &o, &out, &named);
    }
    let request = cli::with_ranges(cli::read_request(input.as_deref())?, &ranges)?;
    let (report, answer) = sheetrs_answer::fill_template_with_answer(
        template,
        program,
        base,
        &request,
        &o,
        Path::new(&out),
    )
    .map_err(|e| Fail::Run(format!("{e:#}")))?;
    log(&wrote_workbook(&named, report.cells));
    println!(
        "{}",
        if pretty {
            serde_json::to_string_pretty(&answer)
        } else {
            serde_json::to_string(&answer)
        }
        .unwrap_or_default()
    );
    Ok(0)
}

/// Run `program` with `request` and write the answer into a copy of `template`.
fn write_workbook(
    program: &Program,
    base: Option<&dyn sheetrs_engine::engine::ConstSource>,
    template: &Book,
    request: &serde_json::Value,
    overrides: &sheetrs_engine::Overrides,
    out: &str,
    named: &str,
) -> Result<i32, Fail> {
    let report =
        sheetrs_answer::fill_template(template, program, base, request, overrides, Path::new(out))
            .map_err(|e| Fail::Run(format!("{e:#}")))?;
    log(&wrote_workbook(named, report.cells));
    Ok(0)
}

/// `-i`, `--input` or `--file`: the workbook (or program) the command works on.
fn input_file(args: &mut Args) -> Result<String, Fail> {
    let mut found = Vec::new();
    for name in ["-i", "--input", "--file"] {
        found.extend(args.values(name)?);
    }
    let mut found = found.into_iter();
    match (found.next(), found.next()) {
        (Some(file), None) => Ok(file),
        (None, _) => Err(Fail::Usage(
            "the workbook is missing: give it with -i FILE (or --input, --file)".into(),
        )),
        _ => Err(Fail::Usage("give one workbook: -i FILE".into())),
    }
}

/// `-o` or `--output`: the file that is written.
fn output_file(args: &mut Args) -> Result<Option<String>, Fail> {
    let mut found = Vec::new();
    for name in ["-o", "--output"] {
        found.extend(args.values(name)?);
    }
    let mut found = found.into_iter();
    match (found.next(), found.next()) {
        (file, None) => Ok(file),
        _ => Err(Fail::Usage("give one output: -o OUT".into())),
    }
}

/// Runs `write` to a file beside `file`, and replaces `file` by it only when it is complete: a failed
/// run leaves `file` as it was and no half-written file behind (the file written is a new one, which
/// the command created and may remove).
fn replace_in_place(
    file: &str,
    write: impl FnOnce(&str) -> Result<i32, Fail>,
) -> Result<i32, Fail> {
    let path = Path::new(file);
    let name = path
        .file_name()
        .and_then(|n| n.to_str())
        .ok_or_else(|| Fail::Run(format!("{file} has no file name")))?;
    let beside = path
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."))
        .join(format!(".{name}.sheetc-{}", std::process::id()));
    let beside_text = beside.to_string_lossy().into_owned();
    let result = write(&beside_text);
    let done = match &result {
        Ok(_) => std::fs::metadata(path)
            .and_then(|m| std::fs::set_permissions(&beside, m.permissions()))
            .and_then(|()| std::fs::rename(&beside, path))
            .map_err(|e| Fail::Run(format!("cannot replace {file}: {e}"))),
        Err(_) => Ok(()),
    };
    if result.is_err() || done.is_err() {
        let _ = std::fs::remove_file(&beside);
    }
    done?;
    result
}

fn load_model(path: &str) -> Result<Model, Fail> {
    Ok(loader::load(
        &Book::from_file(path).map_err(|e| Fail::Run(e.to_string()))?,
    )?)
}

/// Command-line configuration over the file's: `--config`, or `sheetc.toml`
/// next to the workbook.
fn config_from(args: &mut Args) -> Result<(Config, Option<String>), Fail> {
    let explicit = args.value("--config")?;
    let mut cli = Config {
        settings: args.value("--settings")?,
        settings_file: args.value("--settings-file")?.map(PathBuf::from),
        fold: args.flag("--no-fold").then_some(false),
        strict: if args.flag("--no-strict") {
            Some(false)
        } else {
            args.flag("--strict").then_some(true)
        },
        enable_features: args.value("--enable-features")?.map(|v| {
            v.split(',')
                .map(str::trim)
                .filter(|n| !n.is_empty())
                .map(String::from)
                .collect()
        }),
        locale: args.value("--locale")?.map(|tag| LocaleOverride {
            tag: Some(tag),
            ..LocaleOverride::default()
        }),
        ..Config::default()
    };
    let o = parse_overrides(args)?;
    cli.now = o.now;
    cli.seed = o.rand_seed;
    cli.host = Some(o.host);
    Ok((cli, explicit))
}

fn spec_for(
    cli: Config,
    explicit: Option<String>,
    file: &str,
    model: &Model,
) -> Result<Spec, Fail> {
    let path = explicit
        .map(PathBuf::from)
        .or_else(|| config::discover(Path::new(file)));
    let base = match path {
        Some(p) => config::load(&p)?,
        None => Config::default(),
    };
    Ok(cli.over(base).spec(model)?)
}

/// The record of a workbook written with the results of a calculation.
fn wrote_workbook(named: &str, cells: usize) -> Diagnostic {
    Diagnostic::info(
        "WORKBOOK_WRITTEN",
        format!("wrote {named} ({cells} cells written)"),
    )
    .with("path", named)
    .with("cells", cells.to_string())
}

fn print_report(r: &sheetrs_compiler::Report, p: &Program) {
    log(&Diagnostic::info(
        "BUILD_SUMMARY",
        format!(
            "cells: {} total, {} reachable, {} kept | formulas: {} total, {} folded at build time, {} kept | cycles: {}",
            r.cells_total,
            r.cells_reachable,
            r.cells_kept,
            r.formulas_total,
            r.formulas_folded,
            r.formulas_kept,
            r.cycles
        ),
    )
    .with("cells_total", r.cells_total.to_string())
    .with("cells_reachable", r.cells_reachable.to_string())
    .with("cells_kept", r.cells_kept.to_string())
    .with("formulas_total", r.formulas_total.to_string())
    .with("formulas_folded", r.formulas_folded.to_string())
    .with("formulas_kept", r.formulas_kept.to_string())
    .with("cycles", r.cycles.to_string()));
    if r.dynamic_refs {
        log(&Diagnostic::info(
            "BUILD_DYNAMIC_REFERENCES",
            "run-time references (INDIRECT/OFFSET) keep the whole workbook reachable",
        ));
    }
    if !r.unsupported.is_empty() {
        let names = r.unsupported.iter().cloned().collect::<Vec<_>>().join(", ");
        log(&Diagnostic::warning(
            "BUILD_UNKNOWN_NAMES",
            format!("names LibreOffice does not know either (they evaluate to #NAME?): {names}"),
        )
        .with("names", names));
    }
    for w in &r.warnings {
        log(&Diagnostic::warning("BUILD_WARNING", w.clone()));
    }
    for d in p.diagnostics.iter().take(10) {
        log(&Diagnostic::warning("BUILD_DIAGNOSTIC", d.clone()));
    }
    if p.diagnostics.len() > 10 {
        log(&Diagnostic::warning(
            "BUILD_DIAGNOSTICS_OMITTED",
            format!(
                "... {} more diagnostics (see --emit-ir)",
                p.diagnostics.len() - 10
            ),
        )
        .with("omitted", (p.diagnostics.len() - 10).to_string()));
    }
}

/// Stage timing and peak memory on stderr when `SHEETC_TRACE` is set.
fn stage(t: &std::time::Instant, label: &str) {
    if std::env::var_os("SHEETC_TRACE").is_some() {
        eprintln!(
            "[stage {:>8.2?}] {label} (peak {} MB)",
            t.elapsed(),
            sheetrs_compiler::compile::peak_rss_mb().map_or("n/a".to_string(), |mb| mb.to_string())
        );
    }
}

/// An option left over after the known ones were taken is a mistake, said by its name.
fn reject_unknown_options(args: &Args) -> Result<(), Fail> {
    match args.0.iter().find(|a| a.starts_with('-') && a.len() > 1) {
        Some(flag) => Err(Fail::Usage(format!("unknown option `{flag}`"))),
        None => Ok(()),
    }
}

fn build(mut args: Args) -> Result<i32, Fail> {
    let out = output_file(&mut args)?;
    let emit = args.value("--emit-ir")?;
    let runtime = args.value("--runtime")?;
    let wasm = match args.value("--target")?.as_deref() {
        None | Some("native") => false,
        Some("wasm") => true,
        Some(other) => {
            return Err(Fail::Usage(format!(
                "unknown target `{other}` (native or wasm)"
            )));
        }
    };
    let embed_template = args.flag("--embed-template");
    let mode = if embed_template {
        sheetrs_engine::ir::Mode::Embedded
    } else {
        sheetrs_engine::ir::Mode::Values
    };
    let (cli, explicit) = config_from(&mut args)?;
    let file = input_file(&mut args)?;
    let file = file.as_str();
    reject_unknown_options(&args)?;
    if let Some(extra) = args.0.first() {
        return Err(Fail::Usage(format!(
            "unexpected argument `{extra}`: the workbook is given with -i FILE"
        )));
    }
    let t0 = std::time::Instant::now();
    let model = load_model(file)?;
    stage(&t0, "read workbook");
    let mut spec = spec_for(cli, explicit, file, &model)?;
    spec.mode = mode;
    let t1 = std::time::Instant::now();
    let (mut program, report) = compile(model, &spec)?;
    program.build.built_at = sheetrs_compiler::compile::build_time()?;
    stage(&t1, "compile (total)");
    print_report(&report, &program);
    if let Some(p) = emit {
        std::fs::write(&p, serde_json::to_vec_pretty(&program).unwrap_or_default())
            .map_err(|e| Fail::Run(format!("{p}: {e}")))?;
    }
    let stem = Path::new(file)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("model");
    if wasm {
        let out = PathBuf::from(out.unwrap_or_else(|| format!("{stem}.wasm")));
        return write_wasm_module(&mut program, file, embed_template, runtime.as_deref(), &out);
    }
    let out = PathBuf::from(out.unwrap_or_else(|| stem.to_string()));
    let runtime = if let Some(r) = runtime {
        PathBuf::from(r)
    } else {
        let me = std::env::current_exe().map_err(|e| Fail::Run(e.to_string()))?;
        let rt = me.with_file_name(format!("sheetrs-rt{}", std::env::consts::EXE_SUFFIX));
        // The runtime stub cannot write a workbook: an executable that carries its template is
        // made from this binary.
        if rt.exists() && !embed_template {
            rt
        } else {
            me
        }
    };
    let t2 = std::time::Instant::now();
    if embed_template {
        let extension = Path::new(file)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        let bytes = std::fs::read(file).map_err(|e| Fail::Run(format!("{file}: {e}")))?;
        bundle::write_executable_with_template(
            &mut program,
            &runtime,
            &out,
            Some((&extension, &bytes)),
        )?;
    } else {
        bundle::write_executable(&mut program, &runtime, &out)?;
    }
    stage(&t2, "write executable");
    log(
        &Diagnostic::info("EXECUTABLE_WRITTEN", format!("wrote {}", out.display()))
            .with("path", out.display().to_string()),
    );
    // The process ends here: freeing millions of cells one by one would only delay it.
    std::mem::forget(program);
    Ok(0)
}

/// The WebAssembly module of the workbook: the base module and the program (see
/// `sheetrs_compiler::wasm_target`). The base is `--runtime`, or the one next to this binary; there is
/// no other to fall back to.
fn write_wasm_module(
    program: &mut Program,
    file: &str,
    embed_template: bool,
    runtime: Option<&str>,
    out: &Path,
) -> Result<i32, Fail> {
    let base = if let Some(r) = runtime {
        PathBuf::from(r)
    } else {
        let name = if embed_template {
            "sheetrs-rt-template.wasm"
        } else {
            "sheetrs-rt.wasm"
        };
        let me = std::env::current_exe().map_err(|e| Fail::Run(e.to_string()))?;
        let base = me.with_file_name(name);
        if !base.exists() {
            return Err(Fail::Run(format!(
                "the base module `{}` is not next to sheetc: build it with scripts/build-wasm-rt.sh, or give it with --runtime",
                base.display()
            )));
        }
        base
    };
    let base_bytes =
        std::fs::read(&base).map_err(|e| Fail::Run(format!("{}: {e}", base.display())))?;
    let template = if embed_template {
        let extension = Path::new(file)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        let bytes = std::fs::read(file).map_err(|e| Fail::Run(format!("{file}: {e}")))?;
        Some((extension, bytes))
    } else {
        None
    };
    let payload = bundle::payload(
        program,
        template.as_ref().map(|(e, b)| (e.as_str(), b.as_slice())),
    )?;
    let module = sheetrs_compiler::wasm_target::complete(&base_bytes, &payload)?;
    std::fs::write(out, &module).map_err(|e| Fail::Run(format!("{}: {e}", out.display())))?;
    log(&Diagnostic::info(
        "MODULE_WRITTEN",
        format!("wrote {} ({} bytes)", out.display(), module.len()),
    )
    .with("path", out.display().to_string())
    .with("bytes", module.len().to_string()));
    Ok(0)
}

fn calculate(mut args: Args) -> Result<i32, Fail> {
    // The run-time flags go to `calculate_program` untouched; the rest configures the in-memory
    // build and says where the answer goes.
    let mut forwarded: Vec<String> = Vec::new();
    for flag in ["--serve", "--pretty", "--info", "--meta"] {
        if args.flag(flag) {
            forwarded.push(flag.into());
        }
    }
    let cinp = args.value("--request")?;
    if let Some(v) = &cinp {
        forwarded.extend(["--request".into(), v.clone()]);
    }
    for q in args.values("--range")? {
        forwarded.extend(["--range".into(), q]);
    }
    let out = output_file(&mut args)?;
    let in_place = args.flag("--in-place");
    let json = args.flag("--json");
    let (cli, explicit) = config_from(&mut args)?;
    let file = input_file(&mut args)?;
    let file = file.as_str();
    reject_unknown_options(&args)?;
    if let Some(extra) = args.0.first() {
        return Err(Fail::Usage(format!(
            "unexpected argument `{extra}`: the workbook is given with -i FILE"
        )));
    }
    if out.is_some() && in_place {
        return Err(Fail::Usage(
            "-o and --in-place both say where the workbook goes: give one".into(),
        ));
    }
    let writes = out.is_some() || in_place;
    if json && !writes {
        return Err(Fail::Usage(
            "--json adds the JSON to a workbook that is written: give -o OUT or --in-place (without them the answer is JSON already)"
                .into(),
        ));
    }
    let is_json = Path::new(file)
        .extension()
        .is_some_and(|e| e.eq_ignore_ascii_case("json"));
    if writes && is_json {
        return Err(Fail::Usage(
            "a program given as JSON has no workbook to write: give the workbook".into(),
        ));
    }
    if let Some(out) = &out
        && same_file(Path::new(out), Path::new(file))
    {
        return Err(Fail::Usage(format!(
            "the output is the input file ({file}): -o never overwrites the original; use --in-place to replace it"
        )));
    }
    let mode = if writes {
        sheetrs_engine::ir::Mode::Embedded
    } else {
        sheetrs_engine::ir::Mode::Values
    };
    let program: Program = if is_json {
        serde_json::from_slice(&std::fs::read(file).map_err(|e| Fail::Run(format!("{file}: {e}")))?)
            .map_err(|e| Fail::Run(format!("{file}: {e}")))?
    } else {
        let model = load_model(file)?;
        let mut spec = spec_for(cli, explicit, file, &model)?;
        spec.mode = mode;
        let (mut p, report) = compile(model, &spec)?;
        p.build.built_at = sheetrs_compiler::compile::build_time()?;
        print_report(&report, &p);
        p
    };
    if writes {
        let source = Book::from_file(file).map_err(|e| Fail::Run(e.to_string()))?;
        let answer = |target: &str, shown: Option<&str>| {
            let mut again = Vec::new();
            for q in &forwarded {
                again.push(q.clone());
            }
            embedded_answer(
                &program,
                None,
                &source,
                Some(target.to_string()),
                shown,
                json,
                Args(again),
            )
        };
        return if in_place {
            replace_in_place(file, |beside| answer(beside, Some(file)))
        } else {
            answer(out.as_deref().unwrap_or_default(), None)
        };
    }
    Ok(cli::calculate_program(&program, None, Args(forwarded)))
}

/// Whether two paths are the same file (an output that does not exist yet is not the input).
fn same_file(a: &Path, b: &Path) -> bool {
    match (a.canonicalize(), b.canonicalize()) {
        (Ok(a), Ok(b)) => a == b,
        _ => false,
    }
}

fn evaluate(mut args: Args) -> Result<i32, Fail> {
    let flags = cli::parse_model_flags(&mut args)?;
    let tag = args.value("--locale")?;
    let o = parse_overrides(&mut args)?;
    let [formula] = args.0.as_slice() else {
        return Err(Fail::Usage("calc needs exactly one formula".into()));
    };
    let mut m = Model::new();
    let s = m.add_sheet("Sheet1");
    m.set_formula(s, "A1", formula);
    let spec = Spec {
        outputs: vec![sheetrs_compiler::BindingSpec::parse("Sheet1!A1")],
        date_system: flags.date_system,
        iteration: flags.iteration,
        locale: Some(
            sheetrs_compiler::locale::from_tag(
                tag.as_deref().unwrap_or("en"),
                &LocaleOverride::default(),
            )
            // A locale given on the command line that is not valid is a usage error.
            .map_err(|e| Fail::Usage(e.to_string()))?
            .0,
        ),
        now: o.now,
        rand_seed: o.rand_seed,
        host: o.host.clone(),
        ..Spec::default()
    };
    let (program, report) = compile(m, &spec)?;
    for d in &program.diagnostics {
        log(&Diagnostic::warning("BUILD_DIAGNOSTIC", d.clone()));
    }
    let _ = report;
    let out = program.calculate(
        &serde_json::Value::Null,
        &sheetrs_engine::Overrides::default(),
    )?;
    println!("{}", out["Sheet1!A1"]);
    Ok(0)
}

fn inspect(mut args: Args) -> Result<i32, Fail> {
    let top: usize = match args.value("--top")? {
        Some(v) => v
            .parse()
            .map_err(|_| Fail::Usage(format!("bad --top `{v}`")))?,
        None => 30,
    };
    let file = input_file(&mut args)?;
    let file = file.as_str();
    if let Some(extra) = args.0.first() {
        return Err(Fail::Usage(format!(
            "unexpected argument `{extra}`: the workbook is given with -i FILE"
        )));
    }
    let model = load_model(file)?;
    let spec = Spec {
        fold: false,
        inspect: true,
        ..spec_for(
            Config {
                strict: Some(false),
                ..Config::default()
            },
            None,
            file,
            &model,
        )?
    };
    let (program, report) = compile(&model, &spec)?;
    println!("sheets: {}", model.sheets.join(", "));
    println!("date system: {}", model.date_system);
    println!(
        "cells: {} ({} formulas)",
        report.cells_total, report.formulas_total
    );
    println!("circular groups: {}", report.cycles);
    println!(
        "run-time references (INDIRECT/OFFSET): {}",
        if report.dynamic_refs { "yes" } else { "no" }
    );
    if report.unsupported.is_empty() {
        println!("unsupported functions: none");
    } else {
        println!(
            "unsupported functions: {}",
            report
                .unsupported
                .iter()
                .cloned()
                .collect::<Vec<_>>()
                .join(", ")
        );
    }
    let mut uses: Vec<(&String, &usize)> = report.function_uses.iter().collect();
    uses.sort_by(|a, b| b.1.cmp(a.1).then(a.0.cmp(b.0)));
    let total: usize = uses.iter().map(|(_, n)| **n).sum();
    if total > 0 {
        println!("functions used ({} distinct, {total} calls):", uses.len());
        let mut running = 0;
        for (name, n) in uses.iter().take(top) {
            running += **n;
            let status = if report.unsupported.contains(*name) {
                "NOT IMPLEMENTED"
            } else {
                "ok"
            };
            println!(
                "  {name:<18} {n:>8} {:>6.1}%  {status}",
                running as f64 * 100.0 / total as f64
            );
        }
        let missing: usize = uses
            .iter()
            .filter(|(k, _)| report.unsupported.contains(*k))
            .map(|(_, n)| **n)
            .sum();
        println!("  not implemented: {missing} of {total} calls");
    }
    for d in program.diagnostics.iter().take(20) {
        println!("diagnostic: {d}");
    }
    Ok(0)
}

fn verify(mut args: Args) -> Result<i32, Fail> {
    let sheets = args.values("--sheet")?;
    let reference_name = args.value("--reference")?;
    let recalc = match reference_name.as_deref() {
        None | Some("cache") => false,
        Some("libreoffice") => true,
        Some(other) => {
            return Err(Fail::Usage(format!(
                "unknown --reference `{other}` (cache or libreoffice)"
            )));
        }
    };
    let tolerance = match args.value("--tolerance")? {
        Some(v) => v
            .parse::<f64>()
            .map_err(|_| Fail::Usage(format!("bad --tolerance `{v}`")))?,
        None => 1e-9,
    };
    let abs_tolerance = match args.value("--abs-tolerance")? {
        Some(v) => v
            .parse::<f64>()
            .map_err(|_| Fail::Usage(format!("bad --abs-tolerance `{v}`")))?,
        None => 1e-6,
    };
    let max: usize = match args.value("--max")? {
        Some(v) => v
            .parse()
            .map_err(|_| Fail::Usage(format!("bad --max `{v}`")))?,
        None => 20,
    };
    let (mut cli, explicit) = config_from(&mut args)?;
    // The workbook's own settings sheet is irrelevant here, and macros do not
    // change what is being compared.
    cli.strict = cli.strict.or(Some(false));
    let file = input_file(&mut args)?;
    let file = file.as_str();
    reject_unknown_options(&args)?;
    if let Some(extra) = args.0.first() {
        return Err(Fail::Usage(format!(
            "unexpected argument `{extra}`: the workbook is given with -i FILE"
        )));
    }
    let started = std::time::Instant::now();
    let model = load_model(file)?;
    let path = explicit
        .map(PathBuf::from)
        .or_else(|| config::discover(Path::new(file)));
    let base = match path {
        Some(p) => config::load(&p)?,
        None => Config::default(),
    };
    let spec = cli.over(base).base_spec(&model)?;
    // `--reference libreoffice`: the values to compare with come from a LibreOffice recalculation
    // of the file, under the locale the program is built for.
    let reference = if recalc {
        let tag = spec.locale.as_ref().map_or("en-US", |l| l.tag.as_str());
        let dir = std::env::temp_dir().join(format!("sheetc-recalc-{}", std::process::id()));
        let ods = sheetrs_compiler::recalc::recalculate(Path::new(file), tag, &dir)?;
        let loaded = load_model(&ods.to_string_lossy());
        let _ = std::fs::remove_dir_all(&dir);
        Some(loaded?)
    } else {
        None
    };
    let v = sheetrs_compiler::verify::verify(
        &model,
        &spec,
        &sheetrs_compiler::verify::Options {
            sheets,
            tolerance,
            abs_tolerance,
        },
        reference.as_ref(),
    )?;
    println!(
        "{:<28} {:>9} {:>9} {:>9} {:>9} {:>9}",
        "sheet", "formulas", "match", "differ", "volatile", "unparsed"
    );
    for s in v.sheets.iter().filter(|s| s.formulas > 0) {
        println!(
            "{:<28} {:>9} {:>9} {:>9} {:>9} {:>9}",
            s.name.chars().take(28).collect::<String>(),
            s.formulas,
            s.matched,
            s.mismatches.len(),
            s.volatile,
            s.unparsed
        );
    }
    if !v.unsupported.is_empty() {
        println!(
            "unsupported functions: {}",
            v.unsupported.iter().cloned().collect::<Vec<_>>().join(", ")
        );
    }
    let mut kinds: std::collections::BTreeMap<&str, usize> = std::collections::BTreeMap::default();
    for m in v.sheets.iter().flat_map(|s| &s.mismatches) {
        *kinds.entry(m.kind).or_default() += 1;
    }
    if !kinds.is_empty() {
        let list: Vec<String> = kinds.iter().map(|(k, n)| format!("{k}: {n}")).collect();
        println!("differences by kind: {}", list.join(", "));
    }
    for m in v.sheets.iter().flat_map(|s| &s.mismatches).take(max) {
        println!(
            "  {}  [{}] expected {} got {}\n      ={}",
            m.at, m.kind, m.expected, m.got, m.formula
        );
    }
    let total = v.mismatches();
    if total > max {
        println!("  ... {} more (--max N shows more)", total - max);
    }
    log(&Diagnostic::info(
        "VERIFIED",
        format!("verified in {:.1?}", started.elapsed()),
    ));
    Ok(i32::from(total > 0))
}
