//! Prototype of a worksheet reader that parses blocks of rows on several threads.
//!
//! One thread inflates the part and cuts it at row starts; workers parse the blocks into
//! partial sheets that are appended in order. A row without its `r` attribute would lose the
//! count of the rows before it, so the prototype assumes every row has one.

use std::io::Read;
use std::sync::Mutex;
use std::sync::mpsc::{channel, sync_channel};
use std::thread;

use sheetrs_core::Sheet;
use zip::result::ZipError;

use crate::Error;
use crate::worksheet::{ValueContext, parse_worksheet as parse_sequentially};

/// Bytes of inflated XML a block holds, roughly.
const BLOCK_BYTES: usize = 1 << 20;

pub(crate) fn parse_worksheet<R: Read>(
    mut source: R,
    sheet: &mut Sheet,
    context: &ValueContext<'_>,
    part: &str,
    workers: usize,
) -> Result<(), Error> {
    let (job_sender, job_receiver) = sync_channel::<(usize, Vec<u8>)>(workers * 2);
    let job_receiver = Mutex::new(job_receiver);
    let job_receiver = &job_receiver;
    let (result_sender, result_receiver) = channel::<(usize, Result<Sheet, Error>)>();
    thread::scope(move |scope| {
        for _ in 0..workers {
            let result_sender = result_sender.clone();
            scope.spawn(move || {
                loop {
                    let job = job_receiver
                        .lock()
                        .ok()
                        .and_then(|receiver| receiver.recv().ok());
                    let Some((index, bytes)) = job else { break };
                    let mut partial = Sheet::new(String::new());
                    let outcome = parse_sequentially(bytes.as_slice(), &mut partial, context, part)
                        .map(|()| partial);
                    if result_sender.send((index, outcome)).is_err() {
                        break;
                    }
                }
            });
        }
        drop(result_sender);

        let produced = produce_blocks(&mut source, |index, block| {
            job_sender.send((index, block)).is_ok()
        });
        drop(job_sender);

        let mut partials: Vec<(usize, Result<Sheet, Error>)> = result_receiver.iter().collect();
        partials.sort_unstable_by_key(|(index, _)| *index);
        produced?;
        for (_, partial) in partials {
            sheet.append(partial?);
        }
        Ok(())
    })
}

/// Reads `source` in blocks that end where a row starts and hands each to `send`, which
/// returns `false` when nobody is listening any more.
fn produce_blocks<R: Read>(
    source: &mut R,
    mut send: impl FnMut(usize, Vec<u8>) -> bool,
) -> Result<(), Error> {
    let mut carry: Vec<u8> = Vec::new();
    let mut index = 0;
    loop {
        let mut buffer = std::mem::take(&mut carry);
        let kept = buffer.len();
        buffer.resize(kept + BLOCK_BYTES, 0);
        let read = read_fully(source, buffer.get_mut(kept..).unwrap_or_default())
            .map_err(|error| Error::Zip(ZipError::Io(error)))?;
        buffer.truncate(kept + read);
        if read < BLOCK_BYTES {
            send(index, buffer);
            return Ok(());
        }
        match last_row_start(&buffer) {
            Some(cut) if cut > 0 => {
                carry = buffer.split_off(cut);
                if !send(index, buffer) {
                    return Ok(());
                }
                index += 1;
            }
            _ => carry = buffer,
        }
    }
}

fn read_fully<R: Read>(source: &mut R, mut buffer: &mut [u8]) -> std::io::Result<usize> {
    let mut total = 0;
    while !buffer.is_empty() {
        let read = source.read(buffer)?;
        if read == 0 {
            break;
        }
        total += read;
        buffer = buffer.get_mut(read..).unwrap_or_default();
    }
    Ok(total)
}

/// The offset of the last `<row ` or `<row>` in `bytes`.
fn last_row_start(bytes: &[u8]) -> Option<usize> {
    bytes
        .windows(5)
        .rposition(|window| window == b"<row " || window == b"<row>")
}
