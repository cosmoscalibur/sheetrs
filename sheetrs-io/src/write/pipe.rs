//! A part of an archive written while another thread compresses it.
//!
//! The writers rewrite a sheet as a stream, and the archive compresses what they write: both are CPU
//! work (the compression of the output is 13 % of the time of a total recalculation of an XLSX and 21 %
//! of an ODS), and they ran one after the other. Here the thread that rewrites the part hands its output
//! in chunks, through a bounded channel, to a thread that gives it to the archive writer, so the two
//! overlap. Memory is a few chunks, not the part.
//!
//! The bytes the archive receives are the same in the same order, so the file is the same as when
//! one thread did everything. The chunking is the same where there are no threads (WebAssembly), so the
//! deflate stream is the same on every target.

use super::Result;
use std::io::{self, Seek, Write};
#[cfg(not(target_arch = "wasm32"))]
use std::sync::mpsc::{SyncSender, sync_channel};
use zip::ZipWriter;

/// Bytes handed over at a time.
const CHUNK: usize = 256 * 1024;
/// Chunks that wait for the thread that compresses.
#[cfg(not(target_arch = "wasm32"))]
const WAITING: usize = 4;

enum To<'a> {
    #[cfg(not(target_arch = "wasm32"))]
    Thread(SyncSender<Vec<u8>>),
    // Built only where there are no threads (the WebAssembly target); the writes are the same.
    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    Direct(&'a mut dyn Write),
}

/// What the rewriting of a part writes to.
pub(crate) struct Piped<'a> {
    buf: Vec<u8>,
    to: To<'a>,
}

impl Piped<'_> {
    fn send(&mut self) -> io::Result<()> {
        if self.buf.is_empty() {
            return Ok(());
        }
        let chunk = std::mem::replace(&mut self.buf, Vec::with_capacity(CHUNK));
        match &mut self.to {
            #[cfg(not(target_arch = "wasm32"))]
            To::Thread(tx) => tx.send(chunk).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::BrokenPipe,
                    "the archive stopped taking bytes",
                )
            }),
            To::Direct(out) => out.write_all(&chunk),
        }
    }
}

impl Write for Piped<'_> {
    fn write(&mut self, data: &[u8]) -> io::Result<usize> {
        self.buf.extend_from_slice(data);
        if self.buf.len() >= CHUNK {
            self.send()?;
        }
        Ok(data.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Runs `produce`, which writes a part into the sink it is given, with the archive writer compressing in
/// another thread; the part is the one the archive has started (`start_file`). The error of the archive,
/// if there is one, is the one reported, since a failed `produce` is then only a closed channel.
///
/// `produce` runs on the calling thread: it may borrow what it likes.
pub(crate) fn through<W, T>(
    zip: &mut ZipWriter<W>,
    produce: impl FnOnce(&mut Piped<'_>) -> Result<T>,
) -> Result<T>
where
    W: Write + Seek + Send,
{
    #[cfg(not(target_arch = "wasm32"))]
    {
        std::thread::scope(|scope| {
            let (tx, rx) = sync_channel::<Vec<u8>>(WAITING);
            let compressor = scope.spawn(move || -> io::Result<()> {
                let mut result = Ok(());
                for chunk in rx {
                    if result.is_ok() {
                        result = zip.write_all(&chunk);
                    }
                    // After an error the chunks are drained: the producer must not block on a full channel.
                }
                result
            });
            let mut sink = Piped {
                buf: Vec::with_capacity(CHUNK),
                to: To::Thread(tx),
            };
            let produced = produce(&mut sink).and_then(|value| Ok((value, sink.send()?)));
            drop(sink);
            let written = compressor
                .join()
                .unwrap_or_else(|_| Err(io::Error::other("the compressing thread panicked")));
            written?;
            produced.map(|(value, ())| value)
        })
    }
    #[cfg(target_arch = "wasm32")]
    {
        let mut sink = Piped {
            buf: Vec::with_capacity(CHUNK),
            to: To::Direct(zip),
        };
        let value = produce(&mut sink)?;
        sink.send()?;
        Ok(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Read};
    use zip::write::SimpleFileOptions;
    use zip::{ZipArchive, ZipWriter};

    #[test]
    fn a_part_written_through_the_pipe_reads_back_whole() {
        // More than a few chunks, in pieces of an uneven size, as the rewriting of a sheet writes.
        let part: Vec<u8> = (0..3_000_000u32).map(|i| (i % 251) as u8).collect();
        let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
        zip.start_file("part.xml", SimpleFileOptions::default())
            .unwrap();
        through(&mut zip, |out| {
            for piece in part.chunks(7_919) {
                out.write_all(piece)?;
            }
            Ok(())
        })
        .unwrap();
        let bytes = zip.finish().unwrap().into_inner();
        let mut back = Vec::new();
        ZipArchive::new(Cursor::new(bytes))
            .unwrap()
            .by_name("part.xml")
            .unwrap()
            .read_to_end(&mut back)
            .unwrap();
        assert_eq!(back, part);
    }

    /// An output that takes some bytes and then refuses.
    struct Refuses {
        taken: usize,
        inner: Cursor<Vec<u8>>,
    }

    impl Write for Refuses {
        fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
            if self.taken > 100_000 {
                return Err(io::Error::other("the disk is full"));
            }
            self.taken += buf.len();
            self.inner.write(buf)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    impl Seek for Refuses {
        fn seek(&mut self, to: io::SeekFrom) -> io::Result<u64> {
            self.inner.seek(to)
        }
    }

    #[test]
    fn an_output_that_refuses_is_an_error_and_not_a_hang() {
        // The archive is never finished nor dropped: its output refused bytes, so its positions no longer
        // add up and the archive asserts (in a debug build) when it closes. What is tested is the pipe.
        let mut zip = std::mem::ManuallyDrop::new(ZipWriter::new(Refuses {
            taken: 0,
            inner: Cursor::new(Vec::new()),
        }));
        // Incompressible bytes, so that the archive soon writes more than the output takes.
        let mut s = 1u64;
        let noise: Vec<u8> = (0..4_000_000)
            .map(|_| {
                s = s
                    .wrapping_mul(6_364_136_223_846_793_005)
                    .wrapping_add(1_442_695_040_888_963_407);
                (s >> 33) as u8
            })
            .collect();
        zip.start_file("part.xml", SimpleFileOptions::default())
            .unwrap();
        let result = through(&mut zip, |out| {
            for piece in noise.chunks(10_000) {
                out.write_all(piece)?;
            }
            Ok(())
        });
        assert!(result.is_err());
    }
}
