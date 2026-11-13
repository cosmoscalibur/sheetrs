//! The binary encoding of programs and blocks of constants: [`rkyv`] archives, validated on read.
//!
//! An archive is read in place only from memory aligned like the archive itself, and a payload
//! holds its parts at any offset, so [`decode`] copies the bytes it is given into an aligned
//! buffer first; the copy is what the validation reads.

use crate::functions;
use crate::ir::{FuncId, Ns};
use rkyv::api::high::{HighDeserializer, HighSerializer, HighValidator};
use rkyv::bytecheck::CheckBytes;
use rkyv::rancor::{Fallible, Source};
use rkyv::ser::allocator::ArenaHandle;
use rkyv::string::{ArchivedString, StringResolver};
use rkyv::util::AlignedVec;
use rkyv::with::{ArchiveWith, DeserializeWith, SerializeWith};
use rkyv::{Archive, Deserialize, Place, Serialize};
use sheetrs_formula::DateSystem;

/// Encode `value`.
///
/// # Errors
///
/// When the value cannot be encoded.
pub(crate) fn encode<T>(value: &T) -> Result<AlignedVec, rkyv::rancor::Error>
where
    T: for<'a> Serialize<HighSerializer<AlignedVec, ArenaHandle<'a>, rkyv::rancor::Error>>,
{
    rkyv::to_bytes::<rkyv::rancor::Error>(value)
}

/// Decode what [`encode`] wrote, checking that `bytes` is a valid archive of `T` first.
///
/// # Errors
///
/// When `bytes` is not a valid archive of `T`.
pub(crate) fn decode<T>(bytes: &[u8]) -> Result<T, rkyv::rancor::Error>
where
    T: Archive,
    T::Archived: for<'a> CheckBytes<HighValidator<'a, rkyv::rancor::Error>>
        + Deserialize<T, HighDeserializer<rkyv::rancor::Error>>,
{
    let mut aligned = AlignedVec::<16>::with_capacity(bytes.len());
    aligned.extend_from_slice(bytes);
    rkyv::from_bytes::<T, rkyv::rancor::Error>(&aligned)
}

/// An unreadable name inside an archive.
#[derive(Debug, thiserror::Error)]
#[error("{0}")]
struct Unreadable(String);

/// A [`FuncId`] as `ns::NAME`, as in JSON: an index into the registry would not stay valid when
/// the registry grows.
pub(crate) struct Qualified;

impl ArchiveWith<FuncId> for Qualified {
    type Archived = ArchivedString;
    type Resolver = StringResolver;

    fn resolve_with(id: &FuncId, resolver: Self::Resolver, out: Place<Self::Archived>) {
        ArchivedString::resolve_from_str(&id.qualified(), resolver, out);
    }
}

impl<S> SerializeWith<FuncId, S> for Qualified
where
    S: Fallible + rkyv::ser::Writer + ?Sized,
    S::Error: Source,
{
    fn serialize_with(id: &FuncId, s: &mut S) -> Result<Self::Resolver, S::Error> {
        ArchivedString::serialize_from_str(&id.qualified(), s)
    }
}

impl<D> DeserializeWith<ArchivedString, FuncId, D> for Qualified
where
    D: Fallible + ?Sized,
    D::Error: Source,
{
    fn deserialize_with(text: &ArchivedString, _: &mut D) -> Result<FuncId, D::Error> {
        let text = text.as_str();
        text.split_once("::")
            .and_then(|(ns, name)| Ns::parse(ns).map(|ns| (ns, name)))
            .ok_or_else(|| D::Error::new(Unreadable(format!("bad function id `{text}`"))))
            .and_then(|(ns, name)| {
                functions::lookup_in(ns, name)
                    .ok_or_else(|| D::Error::new(Unreadable(format!("unknown function `{text}`"))))
            })
    }
}

/// A JSON value as its text: an archive cannot hold a self-describing value.
pub(crate) struct JsonText;

impl ArchiveWith<serde_json::Value> for JsonText {
    type Archived = ArchivedString;
    type Resolver = StringResolver;

    fn resolve_with(v: &serde_json::Value, resolver: Self::Resolver, out: Place<Self::Archived>) {
        ArchivedString::resolve_from_str(&v.to_string(), resolver, out);
    }
}

impl<S> SerializeWith<serde_json::Value, S> for JsonText
where
    S: Fallible + rkyv::ser::Writer + ?Sized,
    S::Error: Source,
{
    fn serialize_with(v: &serde_json::Value, s: &mut S) -> Result<Self::Resolver, S::Error> {
        ArchivedString::serialize_from_str(&v.to_string(), s)
    }
}

impl<D> DeserializeWith<ArchivedString, serde_json::Value, D> for JsonText
where
    D: Fallible + ?Sized,
    D::Error: Source,
{
    fn deserialize_with(text: &ArchivedString, _: &mut D) -> Result<serde_json::Value, D::Error> {
        serde_json::from_str(text.as_str()).map_err(|e| D::Error::new(Unreadable(e.to_string())))
    }
}

/// [`DateSystem`], which `sheetrs-formula` defines without depending on `rkyv`.
#[derive(Archive, Serialize, Deserialize)]
#[rkyv(remote = DateSystem)]
pub(crate) enum DateSystemDef {
    Null1899,
    Null1900,
    Null1904,
}

impl From<DateSystemDef> for DateSystem {
    fn from(def: DateSystemDef) -> Self {
        match def {
            DateSystemDef::Null1899 => Self::Null1899,
            DateSystemDef::Null1900 => Self::Null1900,
            DateSystemDef::Null1904 => Self::Null1904,
        }
    }
}
