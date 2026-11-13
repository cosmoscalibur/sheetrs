//! Command-line plumbing shared by the compiler binary and the runtime stub.

use crate::Error;
use crate::calculate::Overrides;
use crate::consts::StoredConsts;
use crate::engine::ConstSource;
use crate::ir::Program;
use crate::options::{DateSystem, Iteration};
use std::io::{BufRead, Read, Write};

/// Minimal argument bag: flags may be given as `--k v` or `--k=v`.
#[derive(Debug)]
pub struct Args(pub Vec<String>);

impl Args {
    pub fn flag(&mut self, name: &str) -> bool {
        match self.0.iter().position(|a| a == name) {
            Some(i) => {
                self.0.remove(i);
                true
            }
            None => false,
        }
    }

    /// # Errors
    ///
    /// When the flag is present without a value.
    pub fn value(&mut self, name: &str) -> crate::Result<Option<String>> {
        let eq = format!("{name}=");
        if let Some(i) = self.0.iter().position(|a| a.starts_with(&eq)) {
            return Ok(Some(self.0.remove(i)[eq.len()..].to_string()));
        }
        match self.0.iter().position(|a| a == name) {
            Some(i) if i + 1 < self.0.len() => {
                self.0.remove(i);
                Ok(Some(self.0.remove(i)))
            }
            Some(_) => Err(Error::Usage(format!("{name} needs a value"))),
            None => Ok(None),
        }
    }

    /// # Errors
    ///
    /// When a flag is present without a value.
    pub fn values(&mut self, name: &str) -> crate::Result<Vec<String>> {
        let mut out = Vec::new();
        while let Some(v) = self.value(name)? {
            out.push(v);
        }
        Ok(out)
    }
}

/// `--iterate` (defaults) or `--iterate=N[,EPS]`.
///
/// # Errors
///
/// When an iteration flag has a value that does not parse.
pub fn parse_iteration(args: &mut Args) -> crate::Result<Option<Iteration>> {
    if args.flag("--iterate") {
        return Ok(Some(Iteration::default()));
    }
    let Some(v) = args.value("--iterate")? else {
        return Ok(None);
    };
    let (n, eps) = v.split_once(',').unwrap_or((&v, ""));
    let max_iterations = n
        .parse()
        .map_err(|_| Error::Usage(format!("bad --iterate count `{n}`")))?;
    let max_change = if eps.is_empty() {
        Iteration::default().max_change
    } else {
        eps.parse()
            .map_err(|_| Error::Usage(format!("bad --iterate epsilon `{eps}`")))?
    };
    let it = Iteration {
        max_iterations,
        max_change,
    };
    Ok(Some(it))
}

/// Options for a formula that has no file behind it (`calc`): what a document
/// would otherwise define.
#[derive(Debug, Default)]
pub struct ModelFlags {
    pub date_system: Option<DateSystem>,
    pub iteration: Option<Iteration>,
}

/// # Errors
///
/// When a flag has no value or a value that does not parse.
pub fn parse_model_flags(args: &mut Args) -> crate::Result<ModelFlags> {
    let mut f = ModelFlags::default();
    if let Some(v) = args.value("--date-system")? {
        f.date_system = Some(
            v.parse::<DateSystem>()
                .map_err(|e| Error::Usage(format!("--date-system {e}")))?,
        );
    }
    f.iteration = parse_iteration(args)?;
    Ok(f)
}

/// Flags that a document defines; rejected whenever a file is involved.
pub const FILE_DEFINED_FLAGS: [&str; 4] =
    ["--date-system", "--iterate", "--no-iterate", "--search"];

/// Options of the clock and the random seed, shared by build and run time.
///
/// # Errors
///
/// When a flag has no value or a value that does not parse.
pub fn parse_overrides(args: &mut Args) -> crate::Result<Overrides> {
    for f in FILE_DEFINED_FLAGS {
        if args
            .0
            .iter()
            .any(|a| a == f || a.starts_with(&format!("{f}=")))
        {
            return Err(Error::Usage(format!(
                "{f} is defined by the document (date system, iteration and search settings come from the file); \
                 it is not a command-line option here"
            )));
        }
    }
    let mut o = Overrides::default();
    if args
        .0
        .iter()
        .any(|a| a == "--locale" || a.starts_with("--locale="))
    {
        return Err(Error::Usage(
            "the locale is fixed when the program is built (`--locale` or `[locale]` in the build configuration)"
                .into(),
        ));
    }
    if let Some(v) = args.value("--now")? {
        o.now = Some(
            v.parse()
                .map_err(|_| Error::Usage(format!("bad --now `{v}` (serial number)")))?,
        );
    }
    if let Some(v) = args.value("--seed")? {
        o.rand_seed = Some(
            v.parse()
                .map_err(|_| Error::Usage(format!("bad --seed `{v}`")))?,
        );
    }
    o.host = crate::options::Host {
        system: args.value("--system")?,
        osversion: args.value("--osversion")?,
        release: args.value("--release")?,
        directory: args.value("--directory")?,
        filename: args.value("--filename")?,
    };
    Ok(o)
}

/// The request of a run (`--request`: `inputs`, `ranges`, `majorDimension`): the JSON text itself, `@FILE` for a
/// file that holds it, or `-` for the standard input, which is read only then. Without `--request` there
/// is no request (`null`) and nothing is read: a run that is given no input never waits for one.
///
/// # Errors
///
/// When the file or the standard input cannot be read, or the text is not JSON.
pub fn read_request(value: Option<&str>) -> crate::Result<serde_json::Value> {
    let text = match value {
        None => return Ok(serde_json::Value::Null),
        Some("-") => {
            let mut text = String::new();
            std::io::stdin()
                .read_to_string(&mut text)
                .map_err(Error::Stdin)?;
            text
        }
        Some(v) => match v.strip_prefix('@') {
            Some(path) => std::fs::read_to_string(path).map_err(|e| Error::io(path, e))?,
            None => v.to_string(),
        },
    };
    if text.trim().is_empty() {
        return Ok(serde_json::Value::Null);
    }
    Ok(serde_json::from_str(&text)?)
}

pub const RUNTIME_HELP: &str = "\
Options:
  --request JSON|@FILE|-  the request: a JSON object, given as text, as `@FILE` or as `-` for the
                       standard input (read only then):
                         {\"inputs\": {name: value}, \"ranges\": [\"Sheet1!A1:B3\", name],
                          \"majorDimension\": \"ROWS\"|\"COLUMNS\"}   (all optional)
                       The answer is {\"valueRanges\": [{\"range\", \"name\", \"values\"}...]}:
                       for exactly the `ranges` asked, or, without them, one entry for each
                       output declared when the program was built (`name` is its name).
  --range RANGE        add RANGE to `ranges` (repeatable), like `ranges=` in the Google Sheets
                       API: what may be asked depends on the build (the declared outputs, see --meta)
  --serve              read one request per line, answer one per line
  --pretty             indent the JSON output
  --meta               list the document properties, the settings-sheet entries
                       (inputs, outputs, custom) and the exported defined names
  --info               technical description of the program and how it was built
  --version            the versions of the runtime and of the compiler that built it
  --now SERIAL   --seed N
  (date system, iteration and search settings come from the document; the
  locale is fixed when the program is built)";

/// Add the ranges of the command line to the request's own.
///
/// # Errors
///
/// When the request is not a JSON object or its `ranges` is not an array of strings.
pub fn with_ranges(
    request: serde_json::Value,
    extra: &[String],
) -> crate::Result<serde_json::Value> {
    use serde_json::Value;
    if extra.is_empty() {
        return Ok(request);
    }
    let mut m = match request {
        Value::Null => serde_json::Map::new(),
        Value::Object(m) => m,
        _ => return Err(Error::Request("the request must be a JSON object".into())),
    };
    let mut list = match m.remove("ranges") {
        None => Vec::new(),
        Some(Value::Array(a)) => a,
        Some(_) => {
            return Err(Error::Request(
                "`ranges` must be an array of strings".into(),
            ));
        }
    };
    list.extend(extra.iter().cloned().map(Value::String));
    m.insert("ranges".into(), Value::Array(list));
    Ok(Value::Object(m))
}

fn mode_name(p: &Program) -> &'static str {
    match p.mode {
        crate::ir::Mode::Values => "values",
        crate::ir::Mode::Embedded => "embedded",
        crate::ir::Mode::Recalculate => "recalculate",
    }
}

/// What an executable of this program answers: JSON; the workbook it carries as the template
/// (`--output OUT`), which only an executable that can write workbooks does; or both.
fn answers(p: &Program) -> Vec<&'static str> {
    match p.mode {
        crate::ir::Mode::Values => vec!["json"],
        crate::ir::Mode::Embedded => vec!["json", "workbook"],
        crate::ir::Mode::Recalculate => vec!["workbook"],
    }
}

#[must_use]
pub fn meta_json(p: &Program) -> serde_json::Value {
    serde_json::json!({
        "mode": mode_name(p),
        "answers": answers(p),
        "meta": p.meta,
    })
}

/// Execute a program according to command-line arguments. Returns the exit
/// code: 0 on success, 1 on a failed run, 2 on bad arguments.
#[must_use]
pub fn calculate_program(program: &Program, consts: Option<&StoredConsts>, mut args: Args) -> i32 {
    let base = consts.map(|c| c as &dyn ConstSource);
    let pretty = args.flag("--pretty");
    let serve = args.flag("--serve");
    let info = args.flag("--info");
    if args.flag("--version") {
        println!(
            "sheetrs-engine {} (built by {})",
            env!("CARGO_PKG_VERSION"),
            program.build.compiler
        );
        return 0;
    }
    let meta = args.flag("--meta");
    let input = match args.value("--request") {
        Ok(v) => v,
        Err(e) => return usage_error(&e.to_string()),
    };
    let cli_ranges = match args.values("--range") {
        Ok(v) => v,
        Err(e) => return usage_error(&e.to_string()),
    };
    let overrides = match parse_overrides(&mut args) {
        Ok(o) => o,
        Err(e) => return usage_error(&e.to_string()),
    };
    if let Some(extra) = args.0.first() {
        return usage_error(&format!("unexpected argument `{extra}`\n{RUNTIME_HELP}"));
    }
    let show = |v: &serde_json::Value| {
        if pretty {
            serde_json::to_string_pretty(v)
        } else {
            serde_json::to_string(v)
        }
        .unwrap_or_default()
    };
    if info {
        println!("{}", describe(program, consts));
        return 0;
    }
    if meta {
        println!("{}", show(&meta_json(program)));
        return 0;
    }
    let answer = |j: serde_json::Value| {
        program
            .handle_over(base, &with_ranges(j, &cli_ranges)?, &overrides)
            .map(|value_ranges| value_ranges.to_json())
    };
    if serve {
        let stdin = std::io::stdin();
        let mut out = std::io::stdout().lock();
        for line in stdin.lock().lines() {
            let Ok(line) = line else { break };
            if line.trim().is_empty() {
                continue;
            }
            let reply = serde_json::from_str::<serde_json::Value>(&line)
                .map_err(Error::from)
                .and_then(answer)
                .unwrap_or_else(|e| serde_json::json!({ "error_message": e.to_string() }));
            let _ = writeln!(out, "{}", serde_json::to_string(&reply).unwrap_or_default());
            let _ = out.flush();
        }
        return 0;
    }
    match read_request(input.as_deref()).and_then(answer) {
        Ok(v) => {
            println!("{}", show(&v));
            0
        }
        Err(e) => {
            eprintln!("error: {e}");
            1
        }
    }
}

#[must_use]
pub fn usage_error(msg: &str) -> i32 {
    eprintln!("error: {msg}");
    2
}

/// What `--info` prints: the shape of the program (mode, inputs, outputs, cells, cycles, build).
#[must_use]
pub fn describe(p: &Program, consts: Option<&StoredConsts>) -> String {
    let shape = |b: &crate::ir::Binding| {
        serde_json::json!({
            "name": b.name,
            "range": format!("{}!{}", p.sheets.get(b.rect.sheet as usize).map_or("?", String::as_str), crate::addr::rect_text(&b.rect)),
            "shape": [b.rect.rows(), b.rect.cols()],
        })
    };
    serde_json::to_string_pretty(&serde_json::json!({
        "mode": mode_name(p),
        "answers": answers(p),
        "inputs": p.inputs.iter().map(shape).collect::<Vec<_>>(),
        "outputs": p.outputs.iter().map(shape).collect::<Vec<_>>(),
        "cells": p.cells.len() as u64 + consts.map_or(0, StoredConsts::len),
        "cycles": p.cycles.len(),
        "diagnostics": p.diagnostics,
        "runtime": format!("sheetrs-engine {}", env!("CARGO_PKG_VERSION")),
        "build": p.build,
    }))
    .unwrap_or_default()
}

/// Run `f` on a thread with a large stack: demand-driven evaluation of long
/// dependency chains that the schedule did not order is recursive.
pub fn with_big_stack<T: Send + 'static>(f: impl FnOnce() -> T + Send + 'static) -> T {
    std::thread::Builder::new()
        .stack_size(512 * 1024 * 1024)
        .spawn(f)
        .map_or_else(
            |_| std::process::exit(1),
            |h| h.join().unwrap_or_else(|_| std::process::exit(101)),
        )
}
