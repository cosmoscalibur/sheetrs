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
  sheetc inspect FILE
      Show what the compiler sees: sizes, unsupported functions, cycles.

SETTINGS SHEET
  First row: key, value, description, meta_type. A row with meta_type `input`
  or `output` declares the range in `key` (`Sheet1!B2:C9`); any other type is
  free metadata. The executable lists them with `--meta`. Without the sheet
  every formula range is an output and nothing is an input.

BUILD OPTIONS (each also a key of sheetc.toml; the command line wins)
  --config PATH        configuration file (default: sheetc.toml next to FILE)
  --settings SHEET     settings sheet (default: `settings`, optional)
  --outputs LEVEL      declared (default): only the declared outputs, the most
                       pruning; named: also the workbook's defined names can
                       be queried; any: any range can be queried
  --locale TAG         regional settings, any BCP-47 tag (default `en`); fixed
                       at build time. More in the [locale] table of the file.
  --no-fold            keep every reachable formula (no build-time evaluation)
  --now SERIAL   --seed N
  --runtime PATH       runtime stub to embed into (default: sheetrs-rt next to
                       this binary, else this binary)
  The date system, iteration, search settings and dialect come from the
  document and are not options. `calc` has no document and accepts
  --date-system 1899|1900|1904, --dialect core|xlsx|ods and --iterate[=N[,EPS]].

Executables accept: --input, --serve, --pretty, --meta, --query RANGE, --info,
--now, --seed.";

fn main() {
    let args = Args(std::env::args().skip(1).collect());
    let code = cli::with_big_stack(move || real_main(args));
    std::process::exit(code);
}

fn real_main(mut args: Args) -> i32 {
    if let Some(program) = bundle::read_embedded() {
        return cli::run_program(&program, args);
    }
    if args.0.is_empty() || args.flag("--help") || args.flag("-h") {
        println!("{HELP}");
        return 0;
    }
    let cmd = args.0.remove(0);
    let result = match cmd.as_str() {
        "build" => build(args),
        "run" => run(args),
        "calc" => calc(args),
        "inspect" => inspect(args),
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
        locale: args.value("--locale")?.map(|tag| LocaleOverride {
            tag: Some(tag),
            ..LocaleOverride::default()
        }),
        ..Config::default()
    };
    let o = parse_overrides(args)?;
    cli.now = o.now;
    cli.seed = o.rand_seed;
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
            "warning: unsupported functions (evaluate to #NAME?): {}",
            r.unsupported.iter().cloned().collect::<Vec<_>>().join(", ")
        );
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

fn build(mut args: Args) -> Result<i32, Fail> {
    let out = args.value("-o")?;
    let emit = args.value("--emit-ir")?;
    let runtime = args.value("--runtime")?;
    let (cli, explicit) = config_from(&mut args)?;
    let [file] = args.0.as_slice() else {
        return Err(Fail::Usage("build needs exactly one input file".into()));
    };
    let model = load_model(file)?;
    let spec = spec_for(cli, explicit, file, &model)?;
    let (program, report) = compile(&model, &spec).map_err(|e| Fail::Run(e.0))?;
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
    bundle::write_executable(&program, &runtime, &out).map_err(Fail::Run)?;
    eprintln!("wrote {}", out.display());
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
    for q in args.values("--query")? {
        forwarded.extend(["--query".into(), q]);
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
        let (p, report) = compile(&model, &spec).map_err(|e| Fail::Run(e.0))?;
        print_report(&report, &p);
        p
    };
    Ok(cli::run_program(&program, Args(forwarded)))
}

fn calc(mut args: Args) -> Result<i32, Fail> {
    let flags = cli::parse_model_flags(&mut args)?;
    let tag = args.value("--locale")?;
    let o = parse_overrides(&mut args)?;
    let [formula] = args.0.as_slice() else {
        return Err(Fail::Usage("calc needs exactly one formula".into()));
    };
    let mut m = Model::new(sheetrs_compiler::ir::Ns::Core);
    let s = m.add_sheet("Sheet1");
    m.set_formula(s, "A1", formula);
    let spec = Spec {
        outputs: vec![sheetrs_compiler::BindingSpec::parse("Sheet1!A1")],
        date_system: flags.date_system,
        dialect: flags.dialect,
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

fn inspect(args: Args) -> Result<i32, Fail> {
    let [file] = args.0.as_slice() else {
        return Err(Fail::Usage("inspect needs exactly one input file".into()));
    };
    let model = load_model(file)?;
    let spec = Spec {
        fold: false,
        ..spec_for(Config::default(), None, file, &model)?
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
    for d in program.diagnostics.iter().take(20) {
        println!("diagnostic: {d}");
    }
    Ok(0)
}
