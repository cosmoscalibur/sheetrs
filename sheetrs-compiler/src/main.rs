//! `sheetc`: build, run and inspect spreadsheet programs.

use sheetrs_compiler::config::{self, Config};
use sheetrs_compiler::locale::LocaleOverride;
use sheetrs_compiler::{Model, Spec, compile, loader};
use sheetrs_engine::Program;
use sheetrs_engine::bundle;
use sheetrs_engine::cli::{self, Args, parse_overrides};
use sheetrs_io::reader::Source;
use std::path::{Path, PathBuf};

const HELP: &str = "\
sheetc — compile XLSX/ODS spreadsheets into executables

USAGE
  sheetc build FILE [-o OUT] [options]
      Compile FILE into a standalone executable. The workbook declares its
      inputs, outputs and metadata in a settings sheet (see below).
      `--emit-ir PATH` also writes the program as JSON.
  sheetc run FILE|PROGRAM.json [options] [runtime options]
      Compile in memory and run immediately (JSON on stdin/--input).
  sheetc run FILE --embed-template (--recalculate | inputs declared) --output OUT [options]
      The workbook as its own template (see `--embed-template`): no JSON, the answer is a copy of
      FILE with every formula holding the result of this run and the inputs given in their cells.
  sheetc calc \"=FORMULA\" [--locale TAG] [--now N] [--seed N]
      Evaluate one formula.
  sheetc --version
      The version of the compiler.
  sheetc inspect FILE [--top N]
      Show what the compiler sees: sizes, unsupported functions, cycles, and the
      functions the formulas use, most used first, with whether the engine has them.
  sheetc verify FILE [--sheet NAME]... [--tolerance X] [--abs-tolerance X] [--max N] [--recalc]
                [build options]
      Recompute the formulas of FILE (all sheets, or the named ones) with the
      values the file holds and compare them with the values it cached. Reports
      what differs; exit code 1 if anything does. With --recalc the values to
      compare with come from a full recalculation by LibreOffice ($SOFFICE or
      `soffice`) under the build locale, not from the file's cache.

SETTINGS SHEET
  First row: key, value, description, meta_type. A row with meta_type `input`
  or `output` declares the range in `key` (`Sheet1!B2:C9`); any other type is
  free metadata, and a row without type is the category `extra`. The executable
  lists them with `--meta`, after the properties of the document. Without the sheet
  every formula range is an output and nothing is an input.

BUILD OPTIONS (each also a key of sheetc.toml; the command line wins)
  --config PATH        configuration file (default: sheetc.toml next to FILE)
  --settings SHEET     settings sheet (default: `settings`, optional)
  --settings-file PATH JSON with the same declarations (inputs, outputs, meta), for
                       a workbook without a settings sheet: it compiles and
                       calculates as it is, inputs included. Replaces the sheet
  --outputs LEVEL      declared (default): only the declared outputs, the most
                       pruning; named: also the workbook's defined names can
                       be queried; any: any range can be queried
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
                       this binary, else this binary)
  --embed-template     the workbook is the template of the answer. The executable (made from
                       `sheetc` itself, which writes workbooks) carries it and gives no JSON:
                       `--output OUT` writes a copy of the workbook with the result of every
                       formula and the inputs given (`--input`) in their cells; only the
                       metadata (`--meta`, `--info`) is answered as before. The outputs the
                       workbook declares are discarded (all of it is written); its inputs, if
                       it has any, are what the executable accepts.
  --recalculate        with --embed-template, a workbook that declares no input: the executable
                       recalculates it as it is (required, so that a missing declaration of
                       inputs is not taken for the wish to recalculate)
  The date system, iteration and search settings come from the document and
  are not options. `calc` has no document and accepts
  --date-system 1899|1900|1904|YYYY-MM-DD and --iterate[=N[,EPS]].

Executables accept: --input, --serve, --pretty, --meta, --range RANGE, --info,
--now, --seed, --system, --osversion, --release, --directory, --filename. One built with
--embed-template accepts --output OUT, --input, --meta, --info and the clock and host options.";

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
    let result = match cmd.as_str() {
        "build" => build(args),
        "run" => run(args),
        "calc" => calc(args),
        "inspect" => inspect(args),
        "verify" => verify(args),
        other => Err(Fail::Usage(format!("unknown command `{other}`\n\n{HELP}"))),
    };
    match result {
        Ok(code) => code,
        Err(Fail::Usage(m)) => cli::usage_error(&m),
        Err(Fail::Run(m)) => {
            eprintln!("error: {m}");
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
/// template, the workbook with the answer written into it (`--output OUT`).
fn embedded_workbook(e: &bundle::Embedded, mut args: Args) -> Result<i32, Fail> {
    let Some(template) = &e.template else {
        if args.value("--output")?.is_some() {
            return Err(Fail::Usage(
                "this executable carries no workbook: build it with --embed-template to write one"
                    .into(),
            ));
        }
        return Ok(cli::run_program(&e.program, e.consts.as_ref(), args));
    };
    // The metadata is answered as in any executable; the answer itself is the workbook.
    if args.0.iter().any(|a| a == "--meta" || a == "--info") {
        return Ok(cli::run_program(&e.program, e.consts.as_ref(), args));
    }
    let out = args.value("--output")?.ok_or_else(|| {
        Fail::Usage(
            "this executable writes a workbook, not JSON: give --output OUT (and the inputs with --input)"
                .into(),
        )
    })?;
    let input = args.value("--input")?;
    let o = parse_overrides(&mut args)?;
    if let Some(extra) = args.0.first() {
        return Err(Fail::Usage(format!(
            "unexpected argument `{extra}`: an executable that carries its workbook answers with it"
        )));
    }
    let source = Source::from_bytes(
        template.bytes.clone(),
        template
            .extension
            .parse()
            .map_err(|_| Fail::Run(format!("the template is a `{}`", template.extension)))?,
    );
    write_workbook(
        &e.program,
        e.consts
            .as_ref()
            .map(|c| c as &dyn sheetrs_engine::engine::ConstSource),
        &source,
        input.as_deref(),
        &o,
        &out,
    )
}

/// What the program of an executable that carries its workbook is: everything is computed and
/// written (the outputs of the declarations are discarded) and the inputs are the ones declared; a
/// workbook with none is recalculated as it is, which `--recalculate` says.
fn embed_spec(spec: &mut Spec, recalculate: bool) -> Result<(), Fail> {
    if !spec.outputs.is_empty() {
        eprintln!(
            "note: the outputs the workbook declares are discarded: the executable writes all of it"
        );
        spec.outputs.clear();
    }
    spec.query = sheetrs_compiler::compile::QueryLevel::Any;
    match (spec.inputs.is_empty(), recalculate) {
        (true, false) => Err(Fail::Usage(
            "the workbook declares no input: add them to the settings (sheet or file), or build with --recalculate to recalculate it as it is"
                .into(),
        )),
        (false, true) => Err(Fail::Usage(
            "--recalculate is for a workbook with no input, and this one declares some".into(),
        )),
        _ => Ok(()),
    }
}

/// Run `program` with the request in `input` and write the answer into a copy of `template`.
fn write_workbook(
    program: &Program,
    base: Option<&dyn sheetrs_engine::engine::ConstSource>,
    template: &Source,
    input: Option<&str>,
    overrides: &sheetrs_engine::Overrides,
    out: &str,
) -> Result<i32, Fail> {
    let request = cli::read_input(input)?;
    let report =
        sheetrs_edit::fill_template(template, program, base, &request, overrides, Path::new(out))
            .map_err(|e| Fail::Run(format!("{e:#}")))?;
    eprintln!("wrote {out} ({} cells written)", report.written);
    Ok(0)
}

fn load_model(path: &str) -> Result<Model, Fail> {
    Ok(loader::load(
        &Source::from_file(path).map_err(|e| Fail::Run(e.to_string()))?,
    )?)
}

/// Command-line configuration over the file's: `--config`, or `sheetc.toml`
/// next to the workbook.
fn config_from(args: &mut Args) -> Result<(Config, Option<String>), Fail> {
    let explicit = args.value("--config")?;
    let mut cli = Config {
        settings: args.value("--settings")?,
        settings_file: args.value("--settings-file")?.map(PathBuf::from),
        outputs: args.value("--outputs")?,
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

fn print_report(r: &sheetrs_compiler::Report, p: &Program) {
    eprintln!(
        "cells: {} total, {} reachable, {} kept | formulas: {} total, {} folded at build time, {} kept | cycles: {}",
        r.cells_total,
        r.cells_reachable,
        r.cells_kept,
        r.formulas_total,
        r.formulas_folded,
        r.formulas_kept,
        r.cycles
    );
    if r.dynamic_refs {
        eprintln!("note: run-time references (INDIRECT/OFFSET) keep the whole workbook reachable");
    }
    if !r.unsupported.is_empty() {
        eprintln!(
            "warning: names LibreOffice does not know either (they evaluate to #NAME?): {}",
            r.unsupported.iter().cloned().collect::<Vec<_>>().join(", ")
        );
    }
    for w in &r.warnings {
        eprintln!("warning: {w}");
    }
    for d in p.diagnostics.iter().take(10) {
        eprintln!("diagnostic: {d}");
    }
    if p.diagnostics.len() > 10 {
        eprintln!(
            "... {} more diagnostics (see --emit-ir)",
            p.diagnostics.len() - 10
        );
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

fn build(mut args: Args) -> Result<i32, Fail> {
    let out = args.value("-o")?;
    let emit = args.value("--emit-ir")?;
    let runtime = args.value("--runtime")?;
    let embed_template = args.flag("--embed-template");
    let recalculate = args.flag("--recalculate");
    let (mut cli, explicit) = config_from(&mut args)?;
    if embed_template {
        cli.outputs = Some("any".into());
    } else if recalculate {
        return Err(Fail::Usage("--recalculate needs --embed-template".into()));
    }
    let [file] = args.0.as_slice() else {
        return Err(Fail::Usage("build needs exactly one input file".into()));
    };
    let t0 = std::time::Instant::now();
    let model = load_model(file)?;
    stage(&t0, "read workbook");
    let mut spec = spec_for(cli, explicit, file, &model)?;
    if embed_template {
        embed_spec(&mut spec, recalculate)?;
    }
    let t1 = std::time::Instant::now();
    let (mut program, report) = compile(&model, &spec)?;
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
    eprintln!("wrote {}", out.display());
    // The process ends here: freeing millions of cells one by one would only delay it.
    std::mem::forget(program);
    std::mem::forget(model);
    Ok(0)
}

fn run(mut args: Args) -> Result<i32, Fail> {
    // Run-time flags go to `run_program` untouched; the rest configures the
    // in-memory build.
    let mut forwarded: Vec<String> = Vec::new();
    for flag in ["--serve", "--pretty", "--info", "--meta"] {
        if args.flag(flag) {
            forwarded.push(flag.into());
        }
    }
    if let Some(v) = args.value("--input")? {
        forwarded.extend(["--input".into(), v]);
    }
    for q in args.values("--range")? {
        forwarded.extend(["--range".into(), q]);
    }
    let embed_template = args.flag("--embed-template");
    let recalculate = args.flag("--recalculate");
    let write_to = args.value("--output")?;
    let input_path = forwarded
        .iter()
        .position(|f| f == "--input")
        .and_then(|i| forwarded.get(i + 1))
        .cloned();
    let (mut cli, explicit) = config_from(&mut args)?;
    if embed_template {
        cli.outputs = Some("any".into());
    } else if recalculate || write_to.is_some() {
        return Err(Fail::Usage(
            "--output and --recalculate need --embed-template: only a workbook that is its own template is written"
                .into(),
        ));
    }
    let [file] = args.0.as_slice() else {
        return Err(Fail::Usage("run needs exactly one input file".into()));
    };
    let is_json = Path::new(file)
        .extension()
        .is_some_and(|e| e.eq_ignore_ascii_case("json"));
    if embed_template && is_json {
        return Err(Fail::Usage(
            "a program given as JSON has no workbook to be its template".into(),
        ));
    }
    let program: Program = if std::path::Path::new(file)
        .extension()
        .is_some_and(|e| e.eq_ignore_ascii_case("json"))
    {
        serde_json::from_slice(&std::fs::read(file).map_err(|e| Fail::Run(format!("{file}: {e}")))?)
            .map_err(|e| Fail::Run(format!("{file}: {e}")))?
    } else {
        let model = load_model(file)?;
        let mut spec = spec_for(cli, explicit, file, &model)?;
        if embed_template {
            embed_spec(&mut spec, recalculate)?;
        }
        let (mut p, report) = compile(&model, &spec)?;
        p.build.built_at = sheetrs_compiler::compile::build_time()?;
        print_report(&report, &p);
        p
    };
    let asks_metadata = forwarded.iter().any(|f| f == "--meta" || f == "--info");
    if embed_template && write_to.is_none() && !asks_metadata {
        return Err(Fail::Usage(
            "a workbook that is its own template is written: give --output OUT".into(),
        ));
    }
    if let Some(out) = write_to {
        let source = Source::from_file(file).map_err(|e| Fail::Run(e.to_string()))?;
        return write_workbook(
            &program,
            None,
            &source,
            input_path.as_deref(),
            &sheetrs_engine::Overrides::default(),
            &out,
        );
    }
    Ok(cli::run_program(&program, None, Args(forwarded)))
}

fn calc(mut args: Args) -> Result<i32, Fail> {
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
    let (program, report) = compile(&m, &spec)?;
    for d in &program.diagnostics {
        eprintln!("diagnostic: {d}");
    }
    let _ = report;
    let out = program.execute(
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
    let [file] = args.0.as_slice() else {
        return Err(Fail::Usage("inspect needs exactly one input file".into()));
    };
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
    println!("date system: {:?}", model.date_system);
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
    let recalc = args.flag("--recalc");
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
    let [file] = args.0.as_slice() else {
        return Err(Fail::Usage("verify needs exactly one input file".into()));
    };
    let model = load_model(file)?;
    let path = explicit
        .map(PathBuf::from)
        .or_else(|| config::discover(Path::new(file)));
    let base = match path {
        Some(p) => config::load(&p)?,
        None => Config::default(),
    };
    let spec = cli.over(base).base_spec(&model)?;
    // `--recalc`: the values to compare with come from a LibreOffice recalculation
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
    let started = std::time::Instant::now();
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
    eprintln!("verified in {:.1?}", started.elapsed());
    Ok(i32::from(total > 0))
}
