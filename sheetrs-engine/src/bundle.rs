//! Standalone executables: a runtime binary with the program appended as a
//! payload plus a trailer (`magic`, lengths). No toolchain is needed to
//! "compile" and the artifact is a single file.
//!
//! Layout of the payload: the program in `bincode` with its index of constants, the blocks of
//! constants (see [`crate::consts`]), then the trailer: magic, size of the program, size of the
//! blocks. An executable that carries the workbook it was made from, to be used as the template of
//! the answer, adds the workbook after the blocks (its extension in eight bytes, then its bytes) and
//! ends with the trailer of the second version, which also gives the size of the template.

use crate::Error;
use crate::consts::{self, ConstIndex, StoredConsts};
use crate::ir::Program;
use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::Path;

const MAGIC: &[u8; 8] = b"SHEETRS1";
const TRAILER: usize = 24;
/// The trailer of an executable that carries its template.
const MAGIC_TEMPLATE: &[u8; 8] = b"SHEETRS2";
const TRAILER_TEMPLATE: usize = 32;
/// Bytes the extension of the template takes before its content.
const EXTENSION_BYTES: usize = 8;

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
/// When the runtime cannot be read or the executable cannot be written.
pub fn write_executable(program: &mut Program, runtime: &Path, out: &Path) -> crate::Result<()> {
    write_executable_with_template(program, runtime, out, None)
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
    let (index, blocks) = consts::split(program, consts::MIN_STORED)?;
    let skeleton = bincode::serialize(&PackedRef {
        program,
        consts: &index,
    })?;
    let mut bytes = std::fs::read(runtime).map_err(|e| Error::io(runtime, e))?;
    // A runtime that already carries a payload is stripped first.
    if let Some(len) = payload_len(&bytes) {
        bytes.truncate(bytes.len() - len);
    }
    bytes.extend_from_slice(&skeleton);
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
        bytes.extend_from_slice(&(skeleton.len() as u64).to_le_bytes());
        bytes.extend_from_slice(&(blocks.len() as u64).to_le_bytes());
        bytes.extend_from_slice(&(blob.len() as u64).to_le_bytes());
    } else {
        bytes.extend_from_slice(MAGIC);
        bytes.extend_from_slice(&(skeleton.len() as u64).to_le_bytes());
        bytes.extend_from_slice(&(blocks.len() as u64).to_le_bytes());
    }
    let mut f = File::create(out).map_err(|e| Error::io(out, e))?;
    f.write_all(&bytes).map_err(|e| Error::io(out, e))?;
    drop(f);
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        std::fs::set_permissions(out, std::fs::Permissions::from_mode(0o755))
            .map_err(|e| Error::io(out, e))?;
    }
    Ok(())
}

/// Size of the payload and trailer at the end of `bytes`, if there are any.
fn payload_len(bytes: &[u8]) -> Option<usize> {
    let n = bytes.len();
    let word = |at: usize| -> Option<usize> {
        usize::try_from(u64::from_le_bytes(bytes[at..at + 8].try_into().ok()?)).ok()
    };
    if n >= TRAILER_TEMPLATE
        && &bytes[n - TRAILER_TEMPLATE..n - TRAILER_TEMPLATE + 8] == MAGIC_TEMPLATE
    {
        let total = word(n - 24)?
            .checked_add(word(n - 16)?)?
            .checked_add(word(n - 8)?)?
            .checked_add(TRAILER_TEMPLATE)?;
        return (total <= n).then_some(total);
    }
    if n >= TRAILER && &bytes[n - TRAILER..n - TRAILER + 8] == MAGIC {
        let total = word(n - 16)?
            .checked_add(word(n - 8)?)?
            .checked_add(TRAILER)?;
        return (total <= n).then_some(total);
    }
    None
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

/// The program embedded in the running executable, if any.
#[must_use]
pub fn read_embedded() -> Option<Embedded> {
    let mut f = File::open(std::env::current_exe().ok()?).ok()?;
    let size = f.seek(SeekFrom::End(0)).ok()?;
    let word = |tail: &[u8], at: usize| -> Option<u64> {
        Some(u64::from_le_bytes(tail.get(at..at + 8)?.try_into().ok()?))
    };
    // The trailer of the second version is the longer: look for it first.
    let (skeleton, blocks, template) = if size >= TRAILER_TEMPLATE as u64 {
        f.seek(SeekFrom::End(-i64::try_from(TRAILER_TEMPLATE).ok()?))
            .ok()?;
        let mut tail = [0u8; TRAILER_TEMPLATE];
        f.read_exact(&mut tail).ok()?;
        (&tail[..8] == MAGIC_TEMPLATE)
            .then(|| {
                Some((
                    word(&tail, 8)?,
                    word(&tail, 16)?,
                    word(&tail, 24)?,
                    TRAILER_TEMPLATE as u64,
                ))
            })
            .flatten()
            .map(|(a, b, c, t)| (a, b, Some((c, t))))
    } else {
        None
    }
    .or_else(|| {
        f.seek(SeekFrom::End(-i64::try_from(TRAILER).ok()?)).ok()?;
        let mut tail = [0u8; TRAILER];
        f.read_exact(&mut tail).ok()?;
        (&tail[..8] == MAGIC).then(|| Some((word(&tail, 8)?, word(&tail, 16)?, None)))?
    })?;
    let (template_len, trailer) = template.unwrap_or((0, TRAILER as u64));
    let total = skeleton
        .checked_add(blocks)?
        .checked_add(template_len)?
        .checked_add(trailer)?;
    if total > size {
        return None;
    }
    let start = size - total;
    f.seek(SeekFrom::Start(start)).ok()?;
    let mut buf = vec![0u8; usize::try_from(skeleton).ok()?];
    f.read_exact(&mut buf).ok()?;
    let packed: Packed = bincode::deserialize(&buf).ok()?;
    let template = if template_len >= EXTENSION_BYTES as u64 {
        f.seek(SeekFrom::Start(start + skeleton + blocks)).ok()?;
        let mut blob = vec![0u8; usize::try_from(template_len).ok()?];
        f.read_exact(&mut blob).ok()?;
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
    let consts =
        (packed.consts.cells() > 0).then(|| StoredConsts::new(packed.consts, f, start + skeleton));
    Some(Embedded {
        program: packed.program,
        consts,
        template,
    })
}
