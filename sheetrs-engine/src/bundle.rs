//! Standalone executables: a runtime binary with the program appended as a
//! payload plus a trailer (`magic`, lengths). No toolchain is needed to
//! "compile" and the artifact is a single file.
//!
//! Layout of the payload: the program in `bincode` with its index of constants, the blocks of
//! constants (see [`crate::consts`]), then the trailer: magic, size of the program, size of the
//! blocks.

use crate::consts::{self, ConstIndex, StoredConsts};
use crate::ir::Program;
use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::Path;

const MAGIC: &[u8; 8] = b"SHEETRS1";
const TRAILER: usize = 24;

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
pub fn write_executable(program: &mut Program, runtime: &Path, out: &Path) -> Result<(), String> {
    let (index, blocks) = consts::split(program, consts::MIN_STORED)?;
    let skeleton = bincode::serialize(&PackedRef {
        program,
        consts: &index,
    })
    .map_err(|e| e.to_string())?;
    let mut bytes = std::fs::read(runtime).map_err(|e| format!("{}: {e}", runtime.display()))?;
    // A runtime that already carries a payload is stripped first.
    if let Some(len) = payload_len(&bytes) {
        bytes.truncate(bytes.len() - len);
    }
    bytes.extend_from_slice(&skeleton);
    bytes.extend_from_slice(&blocks);
    bytes.extend_from_slice(MAGIC);
    bytes.extend_from_slice(&(skeleton.len() as u64).to_le_bytes());
    bytes.extend_from_slice(&(blocks.len() as u64).to_le_bytes());
    let mut f = File::create(out).map_err(|e| format!("{}: {e}", out.display()))?;
    f.write_all(&bytes).map_err(|e| e.to_string())?;
    drop(f);
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        std::fs::set_permissions(out, std::fs::Permissions::from_mode(0o755))
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}

/// Size of the payload and trailer at the end of `bytes`, if there are any.
fn payload_len(bytes: &[u8]) -> Option<usize> {
    let n = bytes.len();
    let word = |at: usize| -> Option<usize> {
        usize::try_from(u64::from_le_bytes(bytes[at..at + 8].try_into().ok()?)).ok()
    };
    if n >= TRAILER && &bytes[n - TRAILER..n - TRAILER + 8] == MAGIC {
        let total = word(n - 16)?
            .checked_add(word(n - 8)?)?
            .checked_add(TRAILER)?;
        return (total <= n).then_some(total);
    }
    None
}

/// What an executable carries: the program, and its constants if they are kept apart.
pub struct Embedded {
    pub program: Program,
    pub consts: Option<StoredConsts>,
}

/// The program embedded in the running executable, if any.
pub fn read_embedded() -> Option<Embedded> {
    let mut f = File::open(std::env::current_exe().ok()?).ok()?;
    let size = f.seek(SeekFrom::End(0)).ok()?;
    f.seek(SeekFrom::End(-i64::try_from(TRAILER).ok()?)).ok()?;
    let mut tail = [0u8; TRAILER];
    f.read_exact(&mut tail).ok()?;
    if &tail[..8] != MAGIC {
        return None;
    }
    let skeleton = u64::from_le_bytes(tail[8..16].try_into().ok()?);
    let blocks = u64::from_le_bytes(tail[16..].try_into().ok()?);
    let total = skeleton.checked_add(blocks)?.checked_add(TRAILER as u64)?;
    if total > size {
        return None;
    }
    let start = size - total;
    f.seek(SeekFrom::Start(start)).ok()?;
    let mut buf = vec![0u8; usize::try_from(skeleton).ok()?];
    f.read_exact(&mut buf).ok()?;
    let packed: Packed = bincode::deserialize(&buf).ok()?;
    let consts =
        (packed.consts.cells() > 0).then(|| StoredConsts::new(packed.consts, f, start + skeleton));
    Some(Embedded {
        program: packed.program,
        consts,
    })
}
