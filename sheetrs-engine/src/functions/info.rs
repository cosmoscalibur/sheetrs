//! `CELL` and `INFO`: what the engine can say of a cell, of the file and of the machine. The
//! system, the file's directory and the like are those of the program when it runs, unless the
//! build or the run fixes them (`Host`), as it can the clock. The results are checked against the
//! oracle (`tests/assets/semantics.*`); what depends on the machine, by tests that fix it.

use super::Reg;
use crate::addr::{CellRef, col_name};
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

const DOMAIN: ErrorKind = ErrorKind::Domain;

/// The sheet name as an address writes it; one that needs quotes is not reproduced.
fn sheet_label(name: &str) -> Res<&str> {
    let plain = !name.is_empty()
        && name.chars().all(|c| c.is_ascii_alphanumeric() || c == '_')
        && !name.starts_with(|c: char| c.is_ascii_digit());
    if plain {
        Ok(name)
    } else {
        crate::notes::note(crate::notes::CELL_SHEET_NAME);
        Err(ErrorKind::Unsupported)
    }
}

fn address(e: &Engine, c: CellRef) -> Res<String> {
    let cell = format!("${}${}", col_name(c.col), c.row + 1);
    if c.sheet == e.at.sheet {
        return Ok(cell);
    }
    let sheet = sheet_label(&e.sheet_names()[c.sheet as usize])?.to_string();
    Ok(if e.opts.string_ref_excel {
        format!("{sheet}!{cell}")
    } else {
        format!("${sheet}.{cell}")
    })
}

fn cell(e: &mut Engine, a: &[Value]) -> Res<Value> {
    let kind = e.text(&a[0])?.to_uppercase();
    // The reference is looked at first: a value that is not one is `#REF!`, an error stays.
    let at = match a.get(1) {
        None => e.at,
        Some(Value::Ref(r)) => r.top_left(),
        Some(v) => {
            return Err(match e.scalar(v.clone()) {
                Value::Error(k) => k,
                _ => ErrorKind::Ref,
            });
        }
    };
    let own = a.len() < 2;
    match kind.as_str() {
        "ADDRESS" => Ok(Value::text(address(e, at)?)),
        "COL" => Ok(Value::Number(f64::from(at.col + 1))),
        "ROW" => Ok(Value::Number(f64::from(at.row + 1))),
        "SHEET" => Ok(Value::Number(f64::from(at.sheet) + 1.0)),
        "COORD" => Ok(Value::text(format!(
            "${}:${}${}",
            col_name(u32::from(at.sheet)),
            col_name(at.col),
            at.row + 1
        ))),
        "CONTENTS" => {
            if own {
                return Err(ErrorKind::Circ);
            }
            Ok(match e.cell_value(at) {
                Value::Empty => Value::Number(0.0),
                v => v,
            })
        }
        "TYPE" => {
            if own {
                return Ok(Value::text("v"));
            }
            Ok(Value::text(match e.cell_value(at) {
                Value::Empty => "b",
                Value::Text(_) => "l",
                _ => "v",
            }))
        }
        "FILENAME" => Ok(Value::text(filename(e, at))),
        "COLOR" | "FORMAT" | "PARENTHESES" | "PREFIX" | "PROTECT" | "WIDTH" => {
            crate::notes::note(crate::notes::CELL_FORMAT);
            Err(ErrorKind::Unsupported)
        }
        _ => Err(DOMAIN),
    }
}

/// The build of the LibreOffice whose results the engine reproduces, which `INFO("release")` of
/// that program is.
const RELEASE: &str = "bf6ef2662401c7d6dce671ab4deb26ee0af3a7c1";

/// The name LibreOffice gives the system of the program.
fn system() -> &'static str {
    match std::env::consts::OS {
        "windows" => "WNT",
        "macos" => "MACOSX",
        _ => "LINUX",
    }
}

/// `Linux 6.18`; the other systems give what LibreOffice gives for compatibility with Excel.
fn os_version() -> String {
    if std::env::consts::OS == "linux"
        && let Ok(release) = std::fs::read_to_string("/proc/sys/kernel/osrelease")
    {
        let mut parts = release.trim().split(|c: char| !c.is_ascii_digit());
        if let (Some(major), Some(minor)) = (parts.next(), parts.next()) {
            return format!("Linux {major}.{minor}");
        }
    }
    "Windows (32-bit) NT 5.01".to_string()
}

/// A path as a `file:` URL, as LibreOffice writes the part before the file name.
fn file_url(path: &str) -> String {
    let path = path.replace('\\', "/");
    let mut out = String::from("file://");
    if !path.starts_with('/') {
        out.push('/');
    }
    for b in path.bytes() {
        match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'.' | b'_' | b'~' | b'/' | b':' => {
                out.push(b as char);
            }
            _ => {
                use std::fmt::Write;
                let _ = write!(out, "%{b:02X}");
            }
        }
    }
    out
}

/// `CELL("filename")`: the directory of the program (or the one given), the file name and the
/// sheet, in the syntax of the document.
fn filename(e: &Engine, at: CellRef) -> String {
    let dir = e.opts.host.directory.clone().unwrap_or_else(|| {
        std::env::current_exe()
            .ok()
            .and_then(|p| p.parent().map(|d| d.to_string_lossy().into_owned()))
            .unwrap_or_default()
    });
    let dir = file_url(dir.trim_end_matches(['/', '\\']));
    let name = e.opts.host.filename.clone().unwrap_or_default();
    let sheet = &e.sheet_names()[at.sheet as usize];
    if e.opts.string_ref_excel {
        format!("{dir}/[{name}]{sheet}")
    } else {
        format!("'{dir}/{name}'#${sheet}")
    }
}

fn info(e: &mut Engine, a: &[Value]) -> Res<Value> {
    match e.text(&a[0])?.to_uppercase().as_str() {
        "RECALC" => Ok(Value::text("Automatic")),
        "NUMFILE" => Ok(Value::Number(1.0)),
        "ORIGIN" | "MEMAVAIL" | "TOTMEM" | "MEMUSED" => Err(ErrorKind::Na),
        // LibreOffice answers `#N/A` for the directory too.
        "DIRECTORY" => Err(ErrorKind::Na),
        "SYSTEM" => Ok(Value::text(
            e.opts
                .host
                .system
                .clone()
                .unwrap_or_else(|| system().to_string()),
        )),
        "OSVERSION" => Ok(Value::text(
            e.opts.host.osversion.clone().unwrap_or_else(os_version),
        )),
        "RELEASE" => Ok(Value::text(
            e.opts
                .host
                .release
                .clone()
                .unwrap_or_else(|| RELEASE.to_string()),
        )),
        _ => Err(DOMAIN),
    }
}

pub(super) fn register(r: &mut Reg) {
    // What `INFO` says of the machine is not known when the program is built (nor is the directory
    // of `CELL("filename")`: the compiler keeps those cells out of the folding).
    r.eager("CELL", 1, 2, cell);
    r.eager("INFO", 1, 1, info).volatile();
}
