//! Standalone executables: a runtime binary with the program appended as a
//! payload plus a trailer (`magic`, schema, lengths). No toolchain is needed to
//! "compile" and the artifact is a single file.
//!
//! Layout of the payload: the program in `postcard` with its index of constants, the blocks of
//! constants (see [`crate::consts`]), then the trailer: magic, the schema number, size of the program, size
//! of the blocks. An executable that carries the workbook it was made from, to be used as the template of
//! the answer, adds the workbook after the blocks (its extension in eight bytes, then its bytes) and
//! ends with the trailer of the second kind, which also gives the size of the template.
//!
//! # The schema number
//!
//! The program is serialized with `postcard`, which is positional: a field added, removed or moved in
//! [`Program`] or in what it holds (`BuildOptions`...) changes what the same bytes mean, and a runtime reads
//! them without knowing it, as garbage or as a failure far from the cause. The payload therefore carries
//! [`SCHEMA`], the number of the layout of the program, **next to the magic**, and a runtime refuses a payload
//! whose number is not its own, naming both ([`PayloadError::Schema`]). **Any change of the order or of the type
//! of a field of what is serialized bumps [`SCHEMA`]** (the test `the_schema_number_follows_the_layout_of_the_program`
//! fails until it does). A payload written before the number existed has the number 0.
//!
//! The last bytes of a payload never change their meaning, whatever the schema: the magic (`SHEETRS1`, or
//! `SHEETRS2` with a template), the schema as a little-endian `u32` and four zero bytes, then the lengths. A new
//! layout of those bytes needs a new magic.

use crate::Error;
use crate::consts::{self, ConstIndex, StoredConsts};
use crate::ir::Program;
use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::Path;

/// The number of the layout of the serialized program (the payload of an executable and of a module).
///
/// It is bumped on any change of the order or of the type of a field of [`Program`] or of what it holds (the
/// serialization is positional: see the module). The runtime that reads a payload and the `sheetc` that wrote it have to
/// agree on it, and neither runs a payload of another number.
pub const SCHEMA: u32 = 1;

/// The version of the interface of a WebAssembly module (the exports of `sheetrs-rt-wasm`, `sheetrs_abi`).
///
/// It is an integer bumped **only on an incompatible change of the exports or of what they mean** for a host (a changed
/// signature, a removed export, another meaning of an argument or of a result). An additive export (as
/// `sheetrs_set_calculate_all`) never bumps it: a host finds out whether a module has it by looking for the export.
/// The layout of the payload is [`SCHEMA`], a number of its own.
pub const WASM_ABI_VERSION: u32 = 1;

const MAGIC: &[u8; 8] = b"SHEETRS1";
/// Magic, schema (4 bytes) and four reserved, size of the program, size of the blocks.
const TRAILER: usize = 32;
/// The trailer of an executable that carries its template.
const MAGIC_TEMPLATE: &[u8; 8] = b"SHEETRS2";
const TRAILER_TEMPLATE: usize = 40;
/// The trailers before the schema number existed (schema 0): magic, program, blocks (and template).
const LEGACY_TRAILER: usize = 24;
const LEGACY_TRAILER_TEMPLATE: usize = 32;
/// Bytes the extension of the template takes before its content.
const EXTENSION_BYTES: usize = 8;

/// What a runtime binary says of the payloads it reads: this marker, [`SCHEMA`] as a little-endian `u32`, and [`STUB_END`].
/// `sheetc` looks for it in the stub it appends a program to, and does not write a program a stub cannot read. The end tells the
/// statics from a copy of the marker that the code of the binary holds to compare with.
const STUB_MARKER: &[u8; 16] = b"SHEETRS-STUB-SCH";
const STUB_END: &[u8; 4] = b"SCH!";

/// The marker and the schema every binary that reads payloads holds in its data (the runtime stub and `sheetc`).
#[used]
static STUB_SCHEMA: [u8; 24] = stub_schema();

const fn stub_schema() -> [u8; 24] {
    let mut out = [0u8; 24];
    let schema = SCHEMA.to_le_bytes();
    let mut i = 0;
    while i < 16 {
        out[i] = STUB_MARKER[i];
        i += 1;
    }
    let mut j = 0;
    while j < 4 {
        out[16 + j] = schema[j];
        out[20 + j] = STUB_END[j];
        j += 1;
    }
    out
}

/// Why a payload is not run.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
pub enum PayloadError {
    /// The payload was written for another layout of the program: `found` is its schema number (0 for one written
    /// before the number existed), `expected` the one of this runtime.
    #[error(
        "the program of this executable has payload schema {found}, and this runtime reads schema {expected}: a program is only valid with the `sheetc` that built it (build it again with the `sheetc` of this runtime)"
    )]
    Schema { found: u32, expected: u32 },
    /// The trailer is there and the schema is the runtime's, and the rest does not decode.
    #[error("the program of this executable is damaged")]
    Damaged,
}

#[derive(Serialize)]
struct PackedRef<'a> {
    program: &'a Program,
    consts: &'a ConstIndex,
}

#[derive(Deserialize)]
struct Packed {
    program: Program,
    consts: ConstIndex,
}

/// Write the executable: `runtime` plus `program`. The constants of a big program leave it for the
/// blocks of the payload, so `program` is left without them.
///
/// # Errors
///
/// When the runtime cannot be read, is a stub of another schema, or the executable cannot be written.
pub fn write_executable(program: &mut Program, runtime: &Path, out: &Path) -> crate::Result<()> {
    write_executable_with_template(program, runtime, out, None)
}

/// The schema a runtime binary (a stub, or `sheetc` itself) says it reads, `None` when it says nothing: a runtime of
/// before the schema number existed, or a file that is no runtime.
#[must_use]
pub fn stub_schema_of(runtime: &[u8]) -> Option<u32> {
    let at = runtime.windows(STUB_SCHEMA.len()).position(|window| {
        &window[..STUB_MARKER.len()] == STUB_MARKER && &window[20..] == STUB_END
    })?;
    let bytes = runtime.get(at + STUB_MARKER.len()..at + STUB_MARKER.len() + 4)?;
    Some(u32::from_le_bytes(bytes.try_into().ok()?))
}

/// [`write_executable`] with the workbook the program was made from, as `(extension, bytes)`, to
/// be the template of the answer.
///
/// # Errors
///
/// As [`write_executable`], and when the extension does not fit its eight bytes.
pub fn write_executable_with_template(
    program: &mut Program,
    runtime: &Path,
    out: &Path,
    template: Option<(&str, &[u8])>,
) -> crate::Result<()> {
    let mut bytes = std::fs::read(runtime).map_err(|e| Error::io(runtime, e))?;
    // A runtime that already carries a payload is stripped first.
    if let Some(len) = payload_len(&bytes) {
        bytes.truncate(bytes.len() - len);
    }
    match stub_schema_of(&bytes) {
        Some(schema) if schema == SCHEMA => {}
        Some(schema) => {
            return Err(Error::Usage(format!(
                "the runtime `{}` reads payload schema {schema} and this `sheetc` writes schema {SCHEMA}: a program is only valid with the runtime of the `sheetc` that builds it (use the `sheetrs-rt` of this version)",
                runtime.display()
            )));
        }
        None => {
            return Err(Error::Usage(format!(
                "the runtime `{}` does not say which payload schema it reads (it is of another version of SheetRS, or not a runtime): this `sheetc` writes schema {SCHEMA}",
                runtime.display()
            )));
        }
    }
    let payload = payload(program, template)?;
    bytes.extend_from_slice(&payload);
    File::create(out)
        .and_then(|mut f| f.write_all(&bytes))
        .map_err(|e| Error::io(out, e))?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        std::fs::set_permissions(out, std::fs::Permissions::from_mode(0o755))
            .map_err(|e| Error::io(out, e))?;
    }
    Ok(())
}

/// The payload of a program, as it is appended to a runtime: the program, its blocks of constants,
/// the template workbook if there is one, and the trailer. The constants of a big program leave it
/// for the blocks, so `program` is left without them.
///
/// # Errors
///
/// When the program cannot be encoded, or the extension does not fit its eight bytes.
pub fn payload(program: &mut Program, template: Option<(&str, &[u8])>) -> crate::Result<Vec<u8>> {
    let (index, blocks) = consts::split(program, consts::MIN_STORED)?;
    let skeleton = postcard::to_allocvec(&PackedRef {
        program,
        consts: &index,
    })?;
    let mut bytes = skeleton.clone();
    bytes.extend_from_slice(&blocks);
    if let Some((extension, workbook)) = template {
        if extension.len() > EXTENSION_BYTES || !extension.is_ascii() {
            return Err(Error::Usage(format!(
                "the template's extension `{extension}` is not a short ASCII name"
            )));
        }
        let mut blob = extension.as_bytes().to_vec();
        blob.resize(EXTENSION_BYTES, 0);
        blob.extend_from_slice(workbook);
        bytes.extend_from_slice(&blob);
        bytes.extend_from_slice(MAGIC_TEMPLATE);
        bytes.extend_from_slice(&SCHEMA.to_le_bytes());
        bytes.extend_from_slice(&[0; 4]);
        bytes.extend_from_slice(&(skeleton.len() as u64).to_le_bytes());
        bytes.extend_from_slice(&(blocks.len() as u64).to_le_bytes());
        bytes.extend_from_slice(&(blob.len() as u64).to_le_bytes());
    } else {
        bytes.extend_from_slice(MAGIC);
        bytes.extend_from_slice(&SCHEMA.to_le_bytes());
        bytes.extend_from_slice(&[0; 4]);
        bytes.extend_from_slice(&(skeleton.len() as u64).to_le_bytes());
        bytes.extend_from_slice(&(blocks.len() as u64).to_le_bytes());
    }
    Ok(bytes)
}

/// Where the parts of a payload are, from its trailer.
struct Located {
    /// The schema number of the payload (0: it has none).
    schema: u32,
    skeleton: u64,
    blocks: u64,
    template: u64,
    trailer: u64,
}

impl Located {
    /// Bytes the whole payload takes, trailer included.
    fn total(&self) -> Option<u64> {
        self.skeleton
            .checked_add(self.blocks)?
            .checked_add(self.template)?
            .checked_add(self.trailer)
    }
}

/// The trailer at the end of `tail` (the last bytes of the data, at most the longest trailer): `None` when it ends with
/// no magic. The schema number is read before anything else of a trailer of a layout that may not be this one.
fn locate(tail: &[u8]) -> Option<Located> {
    let n = tail.len();
    let word = |at: usize| -> Option<u64> {
        Some(u64::from_le_bytes(tail.get(at..at + 8)?.try_into().ok()?))
    };
    let schema_at = |at: usize| -> Option<u32> {
        Some(u32::from_le_bytes(
            tail.get(at + 8..at + 12)?.try_into().ok()?,
        ))
    };
    let magic = |from_end: usize, magic: &[u8; 8]| -> Option<usize> {
        let at = n.checked_sub(from_end)?;
        (&tail[at..at + 8] == magic).then_some(at)
    };
    if let Some(at) = magic(TRAILER_TEMPLATE, MAGIC_TEMPLATE) {
        return Some(Located {
            schema: schema_at(at)?,
            skeleton: word(at + 16)?,
            blocks: word(at + 24)?,
            template: word(at + 32)?,
            trailer: TRAILER_TEMPLATE as u64,
        });
    }
    if let Some(at) = magic(TRAILER, MAGIC) {
        return Some(Located {
            schema: schema_at(at)?,
            skeleton: word(at + 16)?,
            blocks: word(at + 24)?,
            template: 0,
            trailer: TRAILER as u64,
        });
    }
    // Before the number existed (schema 0): no schema in the trailer, whose lengths follow the magic.
    if let Some(at) = magic(LEGACY_TRAILER_TEMPLATE, MAGIC_TEMPLATE) {
        return Some(Located {
            schema: 0,
            skeleton: word(at + 8)?,
            blocks: word(at + 16)?,
            template: word(at + 24)?,
            trailer: LEGACY_TRAILER_TEMPLATE as u64,
        });
    }
    if let Some(at) = magic(LEGACY_TRAILER, MAGIC) {
        return Some(Located {
            schema: 0,
            skeleton: word(at + 8)?,
            blocks: word(at + 16)?,
            template: 0,
            trailer: LEGACY_TRAILER as u64,
        });
    }
    None
}

/// Size of the payload and trailer at the end of `bytes`, if there are any (whatever their schema).
fn payload_len(bytes: &[u8]) -> Option<usize> {
    let tail = &bytes[bytes.len().saturating_sub(TRAILER_TEMPLATE)..];
    let total = locate(tail)?.total()?;
    let total = usize::try_from(total).ok()?;
    (total <= bytes.len()).then_some(total)
}

/// What an executable carries: the program, and its constants if they are kept apart.
#[derive(Debug)]
pub struct Embedded {
    pub program: Program,
    pub consts: Option<StoredConsts>,
    /// The workbook the program was made from, when the executable carries it as the template of
    /// the answer.
    pub template: Option<EmbeddedTemplate>,
}

/// A workbook carried by an executable.
#[derive(Debug, Clone)]
pub struct EmbeddedTemplate {
    /// `xlsx` or `ods`.
    pub extension: String,
    pub bytes: Vec<u8>,
}

/// The program embedded in the running executable: `Ok(None)` when it carries none (it is the tool itself).
///
/// # Errors
///
/// When the executable carries a program written for another schema ([`PayloadError::Schema`]), or one that does not decode.
pub fn read_embedded() -> Result<Option<Embedded>, PayloadError> {
    // Keeps the schema this binary reads in the file (`stub_schema_of`), whatever the optimizer does.
    std::hint::black_box(&STUB_SCHEMA);
    let Some(exe) = std::env::current_exe()
        .ok()
        .and_then(|p| File::open(p).ok())
    else {
        return Ok(None);
    };
    read_from_file(exe)
}

fn read_from_file(mut f: File) -> Result<Option<Embedded>, PayloadError> {
    let io = |_| PayloadError::Damaged;
    let size = f.seek(SeekFrom::End(0)).map_err(io)?;
    let take = size.min(TRAILER_TEMPLATE as u64);
    f.seek(SeekFrom::End(
        -i64::try_from(take).map_err(|_| PayloadError::Damaged)?,
    ))
    .map_err(io)?;
    let mut tail = vec![0u8; usize::try_from(take).map_err(|_| PayloadError::Damaged)?];
    f.read_exact(&mut tail).map_err(io)?;
    let Some(located) = locate(&tail) else {
        return Ok(None);
    };
    if located.schema != SCHEMA {
        return Err(PayloadError::Schema {
            found: located.schema,
            expected: SCHEMA,
        });
    }
    let total = located.total().ok_or(PayloadError::Damaged)?;
    if total > size {
        return Err(PayloadError::Damaged);
    }
    let start = size - total;
    let damaged = || PayloadError::Damaged;
    f.seek(SeekFrom::Start(start)).map_err(io)?;
    let mut buf = vec![0u8; usize::try_from(located.skeleton).map_err(|_| damaged())?];
    f.read_exact(&mut buf).map_err(io)?;
    let packed: Packed = postcard::from_bytes(&buf).map_err(|_| damaged())?;
    let template = if located.template >= EXTENSION_BYTES as u64 {
        f.seek(SeekFrom::Start(start + located.skeleton + located.blocks))
            .map_err(io)?;
        let mut blob = vec![0u8; usize::try_from(located.template).map_err(|_| damaged())?];
        f.read_exact(&mut blob).map_err(io)?;
        let extension = String::from_utf8_lossy(&blob[..EXTENSION_BYTES])
            .trim_end_matches('\0')
            .to_string();
        Some(EmbeddedTemplate {
            extension,
            bytes: blob.split_off(EXTENSION_BYTES),
        })
    } else {
        None
    };
    let consts = (packed.consts.cells() > 0)
        .then(|| StoredConsts::new(packed.consts, f, start + located.skeleton));
    Ok(Some(Embedded {
        program: packed.program,
        consts,
        template,
    }))
}

/// What a payload held in memory carries (a module that has its program in its own linear memory):
/// the same layout as [`payload`], read from the end of `bytes`. The constants stay in `bytes`.
///
/// # Errors
///
/// When `bytes` ends with no trailer or with a damaged one ([`PayloadError::Damaged`]), or the payload was written for
/// another schema ([`PayloadError::Schema`]).
pub fn read_payload(bytes: &'static [u8]) -> Result<Embedded, PayloadError> {
    std::hint::black_box(&STUB_SCHEMA);
    let damaged = || PayloadError::Damaged;
    let n = bytes.len();
    let located = locate(&bytes[n.saturating_sub(TRAILER_TEMPLATE)..]).ok_or_else(damaged)?;
    if located.schema != SCHEMA {
        return Err(PayloadError::Schema {
            found: located.schema,
            expected: SCHEMA,
        });
    }
    let total = located.total().ok_or_else(damaged)?;
    let start = usize::try_from(
        u64::try_from(n)
            .map_err(|_| damaged())?
            .checked_sub(total)
            .ok_or_else(damaged)?,
    )
    .map_err(|_| damaged())?;
    let skeleton_end = start
        .checked_add(usize::try_from(located.skeleton).map_err(|_| damaged())?)
        .ok_or_else(damaged)?;
    let packed: Packed = postcard::from_bytes(bytes.get(start..skeleton_end).ok_or_else(damaged)?)
        .map_err(|_| damaged())?;
    let blocks_end = skeleton_end
        .checked_add(usize::try_from(located.blocks).map_err(|_| damaged())?)
        .ok_or_else(damaged)?;
    let template = if located.template >= EXTENSION_BYTES as u64 {
        let template_len = usize::try_from(located.template).map_err(|_| damaged())?;
        let blob = bytes
            .get(blocks_end..blocks_end.checked_add(template_len).ok_or_else(damaged)?)
            .ok_or_else(damaged)?;
        Some(EmbeddedTemplate {
            extension: String::from_utf8_lossy(&blob[..EXTENSION_BYTES])
                .trim_end_matches('\0')
                .to_string(),
            bytes: blob[EXTENSION_BYTES..].to_vec(),
        })
    } else {
        None
    };
    let consts = (packed.consts.cells() > 0)
        .then(|| StoredConsts::from_memory(packed.consts, bytes, skeleton_end as u64));
    Ok(Embedded {
        program: packed.program,
        consts,
        template,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::addr::{CellRef, Rect};
    use crate::ir::{
        Binding, BuildInfo, BuildOptions, CellRun, Mode, Program, Queryable, SourceInfo,
    };
    use crate::options::EngineOptions;
    use crate::value::Value;

    /// A program with a distinct value in every field that is serialized, so that a field that is moved, added, removed or
    /// retyped changes the bytes.
    fn sample() -> Program {
        Program {
            layout: crate::engine::Layout::default(),
            sheets: vec!["S".into()],
            options: EngineOptions::default(),
            cells: vec![],
            hidden_rows: vec![(1, 2, 3)],
            subtotal_cells: vec![CellRef::new(1, 2, 3)],
            cycles: vec![(4, 5)],
            cache: vec![(6, Value::Number(1.5))],
            inputs: vec![],
            outputs: vec![Binding {
                name: "b".into(),
                rect: Rect::new(1, 2, 3, 4, 5),
            }],
            queryable: Queryable::default(),
            mode: Mode::Embedded,
            formula_cells: vec![CellRun {
                sheet: 1,
                row: 2,
                first: 3,
                last: 4,
            }],
            meta: vec![],
            diagnostics: vec!["d".into()],
            build: BuildInfo {
                compiler: "c".into(),
                target: "t".into(),
                built_at: "a".into(),
                options: BuildOptions {
                    outputs: "o".into(),
                    inputs: "i".into(),
                    locale: "l".into(),
                    fold: true,
                    calculate_all: true,
                    force_calculate_all: false,
                    features: vec!["f".into()],
                },
                source: SourceInfo {
                    file: "x".into(),
                    format: "y".into(),
                    sha256: "z".into(),
                },
            },
        }
    }

    fn hex(bytes: &[u8]) -> String {
        use std::fmt::Write as _;
        bytes.iter().fold(String::new(), |mut out, b| {
            let _ = write!(out, "{b:02x}");
            out
        })
    }

    fn leaked(bytes: Vec<u8>) -> &'static [u8] {
        Vec::leak(bytes)
    }

    /// The serialization is positional, so the bytes of a program that sets every field say whether the order or the type
    /// of one changed: the number is bumped, and then this is written again. A field inserted among the others (as
    /// `cache` and `calculate_all` were) changes what every later field is read as, and a runtime of the other layout
    /// would take the payload for its own.
    #[test]
    fn the_schema_number_follows_the_layout_of_the_program() {
        const SCHEMA_OF_THE_BYTES: u32 = 1;
        const BYTES: &str = "010153<options>000101020301010203010405010601000000000000f83f0001016201020304050001010102030400010164016301740161016f0169016c01010001016601780179017a";
        let program = sample();
        // What the engine options hold (the locale tables) changes with the data, not with the layout of the program: the
        // fields around it are what is pinned.
        let bytes = hex(&postcard::to_allocvec(&program).unwrap()).replacen(
            &hex(&postcard::to_allocvec(&program.options).unwrap()),
            "<options>",
            1,
        );
        assert_eq!(
            (SCHEMA, bytes),
            (SCHEMA_OF_THE_BYTES, BYTES.to_string()),
            "the layout of the serialized program changed: bump SCHEMA in bundle.rs and write SCHEMA_OF_THE_BYTES and BYTES again"
        );
    }

    #[test]
    fn a_payload_is_read_back_by_the_runtime_of_its_schema() {
        for template in [None, Some(("xlsx", &b"workbook"[..]))] {
            let mut program = sample();
            let bytes = payload(&mut program, template).unwrap();
            let read = read_payload(leaked(bytes.clone())).unwrap();
            assert_eq!(read.program, program);
            assert_eq!(
                read.template.map(|t| (t.extension, t.bytes)),
                template.map(|(e, b)| (e.to_string(), b.to_vec()))
            );
            assert_eq!(payload_len(&bytes), Some(bytes.len()));
        }
    }

    /// A payload that the number of the schema does not match is refused naming both numbers (the one found, the one read), and
    /// is not read as a program of the layout the runtime knows.
    #[test]
    fn a_payload_of_another_schema_is_refused_naming_both_numbers() {
        for template in [None, Some(("xlsx", &b"workbook"[..]))] {
            let mut bytes = payload(&mut sample(), template).unwrap();
            let trailer = if template.is_some() {
                TRAILER_TEMPLATE
            } else {
                TRAILER
            };
            let at = bytes.len() - trailer + 8;
            bytes[at..at + 4].copy_from_slice(&(SCHEMA + 1).to_le_bytes());
            let error = read_payload(leaked(bytes)).unwrap_err();
            assert_eq!(
                error,
                PayloadError::Schema {
                    found: SCHEMA + 1,
                    expected: SCHEMA
                }
            );
            let text = error.to_string();
            assert!(
                text.contains(&format!("payload schema {}", SCHEMA + 1))
                    && text.contains(&format!("reads schema {SCHEMA}")),
                "{text}"
            );
        }
    }

    /// Payloads that `sheetc` wrote before the number existed (the commit before the redesign of the calculation mode
    /// wrote these two, of `loan.xlsx`, with and without the template) are refused: their `Program` is another layout.
    #[test]
    fn a_payload_written_before_the_schema_number_is_refused_as_schema_0() {
        let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/assets");
        for name in ["payload_schema0.bin", "payload_schema0_template.bin"] {
            let bytes = std::fs::read(dir.join(name)).unwrap();
            assert_eq!(payload_len(&bytes), Some(bytes.len()), "{name}");
            let error = read_payload(leaked(bytes)).unwrap_err();
            assert_eq!(
                error,
                PayloadError::Schema {
                    found: 0,
                    expected: SCHEMA
                },
                "{name}"
            );
            assert!(error.to_string().contains("schema 0"), "{name}");
        }
    }

    #[test]
    fn what_has_no_trailer_is_damaged() {
        assert_eq!(
            read_payload(b"nothing here").unwrap_err(),
            PayloadError::Damaged
        );
        assert_eq!(payload_len(b"nothing here"), None);
        assert_eq!(read_payload(&[]).unwrap_err(), PayloadError::Damaged);
    }

    #[test]
    fn a_runtime_says_which_schema_it_reads() {
        let mut runtime = b"MZ junk".to_vec();
        assert_eq!(stub_schema_of(&runtime), None);
        runtime.extend_from_slice(&STUB_SCHEMA);
        runtime.extend_from_slice(b"more junk");
        assert_eq!(stub_schema_of(&runtime), Some(SCHEMA));
    }

    /// `sheetc` does not write a program into a runtime that reads another schema, or says nothing of it.
    #[test]
    fn a_program_is_not_written_into_a_runtime_of_another_schema() {
        let dir = std::env::temp_dir().join(format!("sheetrs-bundle-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let out = dir.join("out");
        let stub = |name: &str, schema: Option<u32>| {
            let path = dir.join(name);
            let mut bytes = b"stub".to_vec();
            if let Some(schema) = schema {
                bytes.extend_from_slice(STUB_MARKER);
                bytes.extend_from_slice(&schema.to_le_bytes());
                bytes.extend_from_slice(STUB_END);
            }
            std::fs::write(&path, bytes).unwrap();
            path
        };
        let ok = write_executable(&mut sample(), &stub("ok", Some(SCHEMA)), &out);
        assert!(ok.is_ok(), "{ok:?}");
        let other = write_executable(&mut sample(), &stub("other", Some(SCHEMA + 1)), &out)
            .unwrap_err()
            .to_string();
        assert!(
            other.contains(&format!("schema {}", SCHEMA + 1))
                && other.contains(&format!("schema {SCHEMA}")),
            "{other}"
        );
        let silent = write_executable(&mut sample(), &stub("silent", None), &out)
            .unwrap_err()
            .to_string();
        assert!(
            silent.contains("does not say which payload schema"),
            "{silent}"
        );
        std::fs::remove_dir_all(&dir).unwrap();
    }
}
