//! A reader that keeps the bytes the parser consumes while it is switched on, so that the
//! parts of an XML document around the element the model replaces can be kept as they are.

use std::io::{self, BufRead, Read};

pub(crate) struct Tee<R> {
    inner: R,
    on: bool,
    bytes: Vec<u8>,
}

impl<R: BufRead> Tee<R> {
    pub(crate) fn new(inner: R, on: bool) -> Self {
        Self {
            inner,
            on,
            bytes: Vec::new(),
        }
    }

    /// Stops keeping bytes and hands over those kept so far.
    pub(crate) fn pause(&mut self) -> Vec<u8> {
        self.on = false;
        std::mem::take(&mut self.bytes)
    }

    /// Starts keeping the bytes that follow.
    pub(crate) fn resume(&mut self) {
        self.on = true;
    }

    /// The bytes kept since the last pause.
    pub(crate) fn take(&mut self) -> Vec<u8> {
        std::mem::take(&mut self.bytes)
    }
}

impl<R: BufRead> Read for Tee<R> {
    fn read(&mut self, out: &mut [u8]) -> io::Result<usize> {
        let read = self.inner.read(out)?;
        if self.on {
            self.bytes
                .extend_from_slice(out.get(..read).unwrap_or_default());
        }
        Ok(read)
    }
}

impl<R: BufRead> BufRead for Tee<R> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        self.inner.fill_buf()
    }

    fn consume(&mut self, amount: usize) {
        if self.on
            && let Ok(buffer) = self.inner.fill_buf()
        {
            self.bytes
                .extend_from_slice(buffer.get(..amount).unwrap_or(buffer));
        }
        self.inner.consume(amount);
    }
}

/// `head` without the start tag `<name` that ends it.
pub(crate) fn without_start_tag(mut head: Vec<u8>, name: &str) -> Vec<u8> {
    let mut tag = vec![b'<'];
    tag.extend_from_slice(name.as_bytes());
    if let Some(at) = head
        .windows(tag.len())
        .rposition(|window| window == tag.as_slice())
    {
        head.truncate(at);
    }
    head
}
