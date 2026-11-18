//! The modules of a VBA project, listed without reading their code.
//!
//! The project is a compound file whose `VBA/dir` stream lists its modules.

const SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
const FREE: u32 = 0xFFFF_FFFF;
const DIRECTORY_ENTRY_BYTES: usize = 128;
/// The text of a chunk of a compressed container is at most this long.
const CHUNK_TEXT_BYTES: usize = 4096;
/// The most text that a compressed `dir` stream may hold; real ones are a few kilobytes.
const MAX_TEXT_BYTES: usize = 1 << 24;

/// The modules of the project in `bytes`, and the name of the project, as its `dir` stream lists
/// them; `None` when the bytes are not a compound file that holds that stream.
pub(crate) fn project(bytes: &[u8]) -> Option<(String, Vec<String>)> {
    records(&decompress(&stream(bytes, &["VBA", "dir"])?)?)
}

/// The text of a compressed container: a run of chunks, each of at most 4096 bytes of text that
/// is stored as it is or as tokens that copy what came before.
fn decompress(container: &[u8]) -> Option<Vec<u8>> {
    if *container.first()? != 1 {
        return None;
    }
    let mut out = Vec::new();
    let mut at = 1;
    while at < container.len() {
        let header = usize::from(le_u16(container, at)?);
        let end = at.checked_add((header & 0x0FFF) + 3)?.min(container.len());
        let body = container.get(at + 2..end)?;
        at = end;
        if header & 0x8000 == 0 {
            out.extend_from_slice(body);
        } else {
            expand(body, &mut out)?;
        }
        if out.len() > MAX_TEXT_BYTES {
            return None;
        }
    }
    Some(out)
}

/// Appends to `out` the text of a chunk whose tokens are in `body`; a chunk holds at most 4096
/// bytes of text, whatever its tokens say.
fn expand(body: &[u8], out: &mut Vec<u8>) -> Option<()> {
    let chunk_start = out.len();
    let mut position = 0;
    while position < body.len() {
        let flags = *body.get(position)?;
        position += 1;
        for bit in 0..8 {
            if position >= body.len() {
                break;
            }
            if flags & (1 << bit) == 0 {
                if out.len() - chunk_start >= CHUNK_TEXT_BYTES {
                    return None;
                }
                out.push(*body.get(position)?);
                position += 1;
                continue;
            }
            let token = usize::from(le_u16(body, position)?);
            position += 2;
            // The number of bits that say how far back to copy grows with the text.
            let written = out.len() - chunk_start;
            let offset_bits = written.max(2).next_power_of_two().trailing_zeros().max(4);
            let length = (token & (0xFFFF >> offset_bits)) + 3;
            let offset = (token >> (16 - offset_bits)) + 1;
            for _ in 0..length {
                let source = out.len().checked_sub(offset)?;
                if source < chunk_start || out.len() - chunk_start >= CHUNK_TEXT_BYTES {
                    return None;
                }
                out.push(*out.get(source)?);
            }
        }
    }
    Some(())
}

/// The position after the record at `at`, which says its size.
fn skip_record(dir: &[u8], at: usize) -> Option<usize> {
    let size = usize::try_from(le_u32(dir, at + 2)?).ok()?;
    (at + 6).checked_add(size)
}

/// The project name and the module names in the records of a `dir` stream.
fn records(dir: &[u8]) -> Option<(String, Vec<String>)> {
    let mut name = String::new();
    let mut modules: Vec<String> = Vec::new();
    let mut at = 0;
    while at < dir.len() {
        let id = le_u16(dir, at)?;
        let size = usize::try_from(le_u32(dir, at + 2)?).ok()?;
        let end = (at + 6).checked_add(size)?;
        let data = dir.get(at + 6..end)?;
        at = end;
        match id {
            0x0004 => name = String::from_utf8_lossy(data).into_owned(),
            0x0019 => modules.push(String::from_utf8_lossy(data).into_owned()),
            // The name in UTF-16 replaces the one of the module just before it.
            0x0047 => {
                let units: Vec<u16> = (0..data.len() / 2)
                    .filter_map(|unit| le_u16(data, 2 * unit))
                    .collect();
                *modules.last_mut()? = String::from_utf16_lossy(&units);
            }
            // The version record is longer than its size says.
            0x0009 => at += 6 - size.min(6),
            // A control reference goes on after the size it gives.
            0x002F => {
                if le_u16(dir, at)? == 0x0016 {
                    at = skip_record(dir, at)?;
                    at = skip_record(dir, at)?;
                }
                at = skip_record(dir, at)?;
            }
            0x0010 => break,
            _ => {}
        }
    }
    // A module listed twice is one module, and a project without a name is the one called
    // `Standard`.
    let mut seen = std::collections::HashSet::new();
    modules.retain(|module| seen.insert(module.clone()));
    if name.is_empty() {
        "Standard".clone_into(&mut name);
    }
    Some((name, modules))
}

fn le_u16(bytes: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_le_bytes(bytes.get(at..at + 2)?.try_into().ok()?))
}

fn le_u32(bytes: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes(bytes.get(at..at + 4)?.try_into().ok()?))
}

/// The stream at `path`, which is a list of names from the root, as the compound file at
/// `bytes` holds it. Names are compared without regard to case.
fn stream(bytes: &[u8], path: &[&str]) -> Option<Vec<u8>> {
    let file = Compound::new(bytes)?;
    let directory = file.chain(le_u32(bytes, 48)?)?;
    let (entries, _) = directory.as_chunks::<DIRECTORY_ENTRY_BYTES>();
    let entries: Vec<Entry> = entries
        .iter()
        .map(|entry| Entry::new(entry))
        .collect::<Option<_>>()?;
    let mut entry = entries.first()?;
    for name in path {
        entry = sibling(&entries, entry.child, name)?;
    }
    if entry.kind != 2 {
        return None;
    }
    let size = usize::try_from(entry.size).ok()?;
    let data = if size >= file.cutoff {
        file.chain(entry.start)?
    } else {
        let root = entries.first()?;
        Mini::new(&file, root.start, le_u32(bytes, 60)?)?.chain(entry.start)?
    };
    data.get(..size).map(<[u8]>::to_vec)
}

/// The entry called `name` among the siblings whose tree starts at `first`.
fn sibling<'a>(entries: &'a [Entry], first: u32, name: &str) -> Option<&'a Entry> {
    let wanted = name.to_uppercase();
    let mut pending = vec![first];
    let mut visited = 0;
    while let Some(index) = pending.pop() {
        if index == FREE {
            continue;
        }
        visited += 1;
        if visited > entries.len() {
            return None;
        }
        let entry = entries.get(usize::try_from(index).ok()?)?;
        if entry.name.to_uppercase() == wanted {
            return Some(entry);
        }
        pending.push(entry.left);
        pending.push(entry.right);
    }
    None
}

/// The sectors of a compound file and the table that chains them.
struct Compound<'a> {
    bytes: &'a [u8],
    sector_bytes: usize,
    mini_sector_bytes: usize,
    /// The size from which a stream has sectors of its own instead of mini sectors.
    cutoff: usize,
    fat: Vec<u32>,
}

impl<'a> Compound<'a> {
    fn new(bytes: &'a [u8]) -> Option<Self> {
        if bytes.get(..8)? != SIGNATURE {
            return None;
        }
        // Sectors are 512 or 4096 bytes and mini sectors 64.
        let sector_shift = le_u16(bytes, 30).filter(|shift| matches!(shift, 9 | 12))?;
        let mini_shift = le_u16(bytes, 32).filter(|shift| *shift == 6)?;
        let mut file = Self {
            bytes,
            sector_bytes: 1_usize << sector_shift,
            mini_sector_bytes: 1_usize << mini_shift,
            cutoff: usize::try_from(le_u32(bytes, 56)?).ok()?,
            fat: Vec::new(),
        };
        // The sectors of the allocation table are listed in the header and then in a chain.
        let mut listed: Vec<u32> = (0..109)
            .filter_map(|index| le_u32(bytes, 76 + 4 * index))
            .filter(|index| *index != FREE)
            .collect();
        let mut next = le_u32(bytes, 68)?;
        let sectors_in_file = bytes.len() / file.sector_bytes;
        let mut lists = std::collections::HashSet::new();
        while next != END_OF_CHAIN && next != FREE {
            // A list that comes round again, or more sectors than the file has, is a lie.
            if !lists.insert(next) || listed.len() > sectors_in_file {
                return None;
            }
            let list = file.sector(next)?;
            let entries = file.sector_bytes / 4 - 1;
            listed.extend(
                (0..entries)
                    .filter_map(|index| le_u32(list, 4 * index))
                    .filter(|index| *index != FREE),
            );
            next = le_u32(list, 4 * entries)?;
        }
        if listed.len() > sectors_in_file {
            return None;
        }
        for index in listed {
            let data = file.sector(index)?;
            let entries: Vec<u32> = (0..file.sector_bytes / 4)
                .filter_map(|entry| le_u32(data, 4 * entry))
                .collect();
            file.fat.extend(entries);
        }
        Some(file)
    }

    fn sector(&self, index: u32) -> Option<&'a [u8]> {
        let start = (usize::try_from(index).ok()? + 1).checked_mul(self.sector_bytes)?;
        self.bytes.get(start..start.checked_add(self.sector_bytes)?)
    }

    /// The bytes of the sectors that follow one another from `first`.
    fn chain(&self, first: u32) -> Option<Vec<u8>> {
        let mut data = Vec::new();
        let mut current = first;
        let mut steps = 0;
        while current != END_OF_CHAIN && current != FREE {
            data.extend_from_slice(self.sector(current)?);
            current = *self.fat.get(usize::try_from(current).ok()?)?;
            steps += 1;
            if steps > self.fat.len() {
                return None;
            }
        }
        Some(data)
    }
}

/// The mini stream, which holds the streams too small for sectors of their own.
struct Mini {
    stream: Vec<u8>,
    sector_bytes: usize,
    fat: Vec<u32>,
}

impl Mini {
    fn new(file: &Compound<'_>, start: u32, first_fat_sector: u32) -> Option<Self> {
        let table = file.chain(first_fat_sector)?;
        Some(Self {
            stream: file.chain(start)?,
            sector_bytes: file.mini_sector_bytes,
            fat: (0..table.len() / 4)
                .filter_map(|entry| le_u32(&table, 4 * entry))
                .collect(),
        })
    }

    fn chain(&self, first: u32) -> Option<Vec<u8>> {
        let mut data = Vec::new();
        let mut current = first;
        let mut steps = 0;
        while current != END_OF_CHAIN && current != FREE {
            let start = usize::try_from(current)
                .ok()?
                .checked_mul(self.sector_bytes)?;
            data.extend_from_slice(
                self.stream
                    .get(start..start.checked_add(self.sector_bytes)?)?,
            );
            current = *self.fat.get(usize::try_from(current).ok()?)?;
            steps += 1;
            if steps > self.fat.len() {
                return None;
            }
        }
        Some(data)
    }
}

/// An entry of the directory of a compound file.
struct Entry {
    name: String,
    kind: u8,
    left: u32,
    right: u32,
    child: u32,
    start: u32,
    size: u64,
}

impl Entry {
    fn new(bytes: &[u8]) -> Option<Self> {
        let length = usize::from(le_u16(bytes, 64)?).saturating_sub(2).min(64);
        let units: Vec<u16> = (0..length / 2)
            .filter_map(|unit| le_u16(bytes, 2 * unit))
            .collect();
        Some(Self {
            name: String::from_utf16_lossy(&units),
            kind: *bytes.get(66)?,
            left: le_u32(bytes, 68)?,
            right: le_u32(bytes, 72)?,
            child: le_u32(bytes, 76)?,
            start: le_u32(bytes, 116)?,
            size: u64::from(le_u32(bytes, 120)?),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_token_copies_what_was_written_before() {
        // `abc` as literals, then a copy of six bytes from three back.
        let container = [0x01, 0x05, 0xB0, 0x08, b'a', b'b', b'c', 0x03, 0x20];
        assert_eq!(decompress(&container), Some(b"abcabcabc".to_vec()));
    }

    #[test]
    fn a_chunk_that_is_not_compressed_is_copied() {
        let container = [0x01, 0x02, 0x30, b'a', b'b', b'c'];
        assert_eq!(decompress(&container), Some(b"abc".to_vec()));
    }

    #[test]
    fn a_token_that_would_grow_a_chunk_past_4096_bytes_is_refused() {
        // `a`, then copies of 4098 bytes from one back.
        let mut container = vec![0x01, 0x0F, 0xB0, 0xFE, b'a'];
        for _ in 0..7 {
            container.extend_from_slice(&[0xFF, 0x0F]);
        }
        assert_eq!(decompress(&container), None);
    }

    #[test]
    fn a_header_with_sectors_of_an_impossible_size_is_refused() {
        let mut file = vec![0_u8; 1024];
        file[..8].copy_from_slice(&SIGNATURE);
        file[30] = 63;
        file[32] = 6;
        assert_eq!(project(&file), None);
        file[30] = 9;
        file[32] = 63;
        assert_eq!(project(&file), None);
        file[30] = 0;
        file[32] = 6;
        assert_eq!(project(&file), None);
    }

    #[test]
    fn a_list_of_allocation_sectors_that_comes_round_again_is_refused() {
        let mut file = vec![0_u8; 1024];
        file[..8].copy_from_slice(&SIGNATURE);
        file[30] = 9;
        file[32] = 6;
        // The first list is the sector 0, whose last entry names the sector 0 again.
        file[68..72].copy_from_slice(&0_u32.to_le_bytes());
        file[1020..1024].copy_from_slice(&0_u32.to_le_bytes());
        assert_eq!(project(&file), None);
    }

    #[test]
    fn text_that_is_not_a_compressed_container_has_no_records() {
        assert_eq!(decompress(b"abc"), None);
        assert_eq!(project(b"not a compound file"), None);
    }
}
