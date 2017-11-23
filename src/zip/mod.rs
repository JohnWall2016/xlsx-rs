//! A basic ZipReader/Writer crate

//#![warn(missing_docs)]

mod spec;
mod crc32;
mod types;
pub mod read;
mod compression;
pub mod write;
mod cp437;
pub mod result;

pub use self::read::ZipArchive;
pub use self::write::ZipWriter;
pub use self::compression::CompressionMethod;
