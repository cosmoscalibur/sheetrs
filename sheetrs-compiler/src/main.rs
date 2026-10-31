//! `sheetc`: build, run and inspect spreadsheet programs.

use sheetrs_compiler::cli::{self, Args, parse_overrides};
use sheetrs_compiler::config::{self, Config};
use sheetrs_compiler::locale::LocaleOverride;
use sheetrs_compiler::{Model, Program, Spec, bundle, compile, loader};
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
  The date system, iteration and search settings come from the document and
  are not options. `calc` has no document and accepts
  --date-system 1899|1900|1904|YYYY-MM-DD and --iterate[=N[,EPS]].

Executables accept: --input, --serve, --pretty, --meta, --range RANGE, --info,
--now, --seed, --system, --osversion, --release, --directory, --filename.";

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
        return cli::run_program(&e.program, e.consts.as_ref(), args);
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

fn load_model(path: &str) -> Result<Model, Fail> {
    loader::load(Path::new(path)).map_err(Fail::Run)
}

/// Command-line configuration over the file's: `--config`, or `sheetc.toml`
/// next to the workbook.
fn config_from(args: &mut Args) -> Result<(Config, Option<String>), Fail> {
    let explicit = args.value("--config")?;
    let mut cli = Config {
        settings: args.value("--settings")?,
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
        Some(p) => config::load(&p).map_err(Fail::Run)?,
        None => Config::default(),
    };
    cli.over(base).spec(model).map_err(|e| Fail::Run(e.0))
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
    let (cli, explicit) = config_from(&mut args)?;
    let [file] = args.0.as_slice() else {
        return Err(Fail::Usage("build needs exactly one input file".into()));
    };
    let t0 = std::time::Instant::now();
    let model = load_model(file)?;
    stage(&t0, "read workbook");
    let spec = spec_for(cli, explicit, file, &model)?;
    let t1 = std::time::Instant::now();
    let (mut program, report) = compile(&model, &spec).map_err(|e| Fail::Run(e.0))?;
    program.build.built_at = sheetrs_compiler::compile::build_time().map_err(Fail::Run)?;
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
    let runtime = match runtime {
        Some(r) => PathBuf::from(r),
        None => {
            let me = std::env::current_exe().map_err(|e| Fail::Run(e.to_string()))?;
            let rt = me.with_file_name(format!("sheetrs-rt{}", std::env::consts::EXE_SUFFIX));
            if rt.exists() { rt } else { me }
        }
    };
    let t2 = std::time::Instant::now();
    bundle::write_executable(&mut program, &runtime, &out).map_err(Fail::Run)?;
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
    let (cli, explicit) = config_from(&mut args)?;
    let [file] = args.0.as_slice() else {
        return Err(Fail::Usage("run needs exactly one input file".into()));
    };
    let program: Program = if file.ends_with(".json") {
        serde_json::from_slice(&std::fs::read(file).map_err(|e| Fail::Run(format!("{file}: {e}")))?)
            .map_err(|e| Fail::Run(format!("{file}: {e}")))?
    } else {
        let model = load_model(file)?;
        let spec = spec_for(cli, explicit, file, &model)?;
        let (mut p, report) = compile(&model, &spec).map_err(|e| Fail::Run(e.0))?;
        p.build.built_at = sheetrs_compiler::compile::build_time().map_err(Fail::Run)?;
        print_report(&report, &p);
        p
    };
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
            )?
            .0,
        ),
        now: o.now,
        rand_seed: o.rand_seed,
        host: o.host.clone(),
        ..Spec::default()
    };
    let (program, report) = compile(&m, &spec).map_err(|e| Fail::Run(e.0))?;
    for d in &program.diagnostics {
        eprintln!("diagnostic: {d}");
    }
    let _ = report;
    let out = program
        .execute(
            &serde_json::Value::Null,
            &sheetrs_compiler::Overrides::default(),
        )
        .map_err(|e| Fail::Run(e.0))?;
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
    let (program, report) = compile(&model, &spec).map_err(|e| Fail::Run(e.0))?;
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
        Some(p) => config::load(&p).map_err(Fail::Run)?,
        None => Config::default(),
    };
    let spec = cli
        .over(base)
        .base_spec(&model)
        .map_err(|e| Fail::Run(e.0))?;
    // `--recalc`: the values to compare with come from a LibreOffice recalculation
    // of the file, under the locale the program is built for.
    let reference = if recalc {
        let tag = spec.locale.as_ref().map_or("en-US", |l| l.tag.as_str());
        let dir = std::env::temp_dir().join(format!("sheetc-recalc-{}", std::process::id()));
        let ods =
            sheetrs_compiler::recalc::recalculate(Path::new(file), tag, &dir).map_err(Fail::Run)?;
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
    )
    .map_err(|e| Fail::Run(e.0))?;
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
    let mut kinds: std::collections::BTreeMap<&str, usize> = Default::default();
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
