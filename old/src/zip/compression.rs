//! Possible ZIP compression methods.

use std::fmt;

/// Compression methods for the contents of a ZIP file.
#[derive(Copy, Clone, PartialEq, Debug)]
pub enum CompressionMethod
{
    /// The file is stored (no compression)
    Stored,
    /// The file is Deflated
    Deflated,
    /// File is compressed using BZIP2 algorithm
    #[cfg(feature = "bzip2")]
    Bzip2,
    /// Unsupported compression method
    Unsupported(u16),
}

impl CompressionMethod {
    /// Converts an u16 to its corresponding CompressionMethod
    pub fn from_u16(val: u16) -> CompressionMethod {
        match val {
            0 => CompressionMethod::Stored,
            8 => CompressionMethod::Deflated,
            #[cfg(feature = "bzip2")]
            12 => CompressionMethod::Bzip2,
            v => CompressionMethod::Unsupported(v),
        }
    }

    /// Converts a CompressionMethod to a u16
    pub fn to_u16(self) -> u16 {
        match self {
            CompressionMethod::Stored => 0,
            CompressionMethod::Deflated => 8,
            #[cfg(feature = "bzip2")]
            CompressionMethod::Bzip2 => 12,
            CompressionMethod::Unsupported(v) => v,
        }
    }
}

impl fmt::Display for CompressionMethod {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        // Just duplicate what the Debug format looks like, i.e, the enum key:
        write!(f, "{:?}", self)
    }
}
