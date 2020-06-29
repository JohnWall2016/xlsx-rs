#![allow(dead_code)]

use std::io::Read;
use std::error::Error;

use yaserde::{de, ser, YaDeserialize, YaSerialize};

type XlsxResult<T> = std::result::Result<T, Box<dyn Error>>;

trait ArchiveDeserable : Sized {
    fn archive_string(ar: &mut zip::Archive) -> XlsxResult<String> {
        use self::zip::ReadAll;
        Ok(ar.by_name(Self::path())?.read_all_to_string()?)
    }

    fn archive_reader(ar: &mut zip::Archive) -> XlsxResult<::zip::read::ZipFile> {
        Ok(ar.by_name(Self::path())?)
    }

    fn path() -> &'static str;

    fn from_archive(ar: &mut zip::Archive) -> XlsxResult<Self>;

    fn to_string(&self) -> XlsxResult<String>;
}

trait YaDeserable: Sized {
    fn from_str(s: &str) -> XlsxResult<Self>;

    fn from_reader<R: Read>(reader: R) -> XlsxResult<Self>;

    fn to_string(&self) -> XlsxResult<String>;
}

impl<T: YaSerialize + YaDeserialize> YaDeserable for T {
    fn from_str(s: &str) -> XlsxResult<T> {
        Ok(de::from_str(s)?)
    }

    fn from_reader<R: Read>(reader: R) -> XlsxResult<T> {
        Ok(de::from_reader(reader)?)
    }

    fn to_string(&self) -> XlsxResult<String> {
        Ok(ser::to_string(self)?)
    }
}

#[macro_export]
macro_rules! ar_deserable {
    ($type: ident, $path: expr, $field: ident: $field_type: ty) => {
        use crate::xlsx::zip;
        use yaserde::de::from_reader;
        use yaserde::ser::to_string;

        impl ArchiveDeserable for $type {
            fn path() -> &'static str {
                $path
            }
        
            fn from_archive(ar: &mut zip::Archive) -> XlsxResult<Self> {
                Ok($type {
                    $field: from_reader(ar.by_name(Self::path())?)?
                })
            }

            fn to_string(&self) -> XlsxResult<String> {
                Ok(to_string(&self.$field)?)
            }
        }
    }
}

#[macro_export]
macro_rules! enum_default {
    ($type: ident :: $variant: ident) => {
        impl Default for $type {
            fn default() -> Self {
                Self::$variant
            }
        }
    }
}

mod zip;
mod content_types;
mod app_properties;
mod core_properties;
mod relationships;
mod shared_strings;
mod style_sheet;
mod worksheet;
mod row;

mod workbook;

#[cfg(test)]
mod test {
    use super::*;

    pub fn test_file() -> String {
        return format!("{}/test_data/table.xlsx", env!("CARGO_MANIFEST_DIR"));
    }

    pub fn test_archive() -> XlsxResult<zip::Archive> {
        Ok(zip::Archive::new(test_file())?)
    }
}