#![allow(dead_code)]

use yaserde::de::{from_reader, from_str};
use yaserde::ser::to_string;
use yaserde::{YaDeserialize, YaSerialize};

use std::error::Error;

type XlsxResult<T> = std::result::Result<T, Box<dyn Error>>;

trait ArchiveDeserable<D: YaDeserialize, S: YaSerialize = D>: Sized {
    fn path() -> &'static str;
    
    fn deseralize_to(de: D) -> XlsxResult<Self>;

    fn seralize_to(&self) -> XlsxResult<&S>;

    fn load_archive(ar: &mut zip::Archive) -> XlsxResult<Self> {
        Self::deseralize_to(from_reader(ar.by_name(Self::path())?)?)
    }

    fn load_string(s: &str) -> XlsxResult<Self> {
        Self::deseralize_to(from_str(s)?)
    }

    fn archive_str(ar: &mut zip::Archive) -> XlsxResult<String> {
        use self::zip::ReadAll;
        Ok(ar.by_name(Self::path())?.read_all_to_string()?)
    }

    fn to_string(&self) -> XlsxResult<String> {
        Ok(to_string(self.seralize_to()?)?)
    }
}

#[macro_export]
macro_rules! ar_deserable {
    ($type: ident, $path: expr, $field: ident: $field_type: ty, $closure: expr) => {
        impl ArchiveDeserable<$field_type> for $type {
            fn path() -> &'static str {
                $path
            }
        
            fn deseralize_to(de: $field_type) -> XlsxResult<Self> {
                Ok($closure(de))
            }
        
            fn seralize_to(&self) -> XlsxResult<&$field_type> {
                Ok(&self.$field)
            }
        }
    };
    ($type: ident, $path: expr, $field: ident: $field_type: ty) => {
        impl ArchiveDeserable<$field_type> for $type {
            fn path() -> &'static str {
                $path
            }
        
            fn deseralize_to(de: $field_type) -> XlsxResult<Self> {
                Ok($type{ $field: de })
            }
        
            fn seralize_to(&self) -> XlsxResult<&$field_type> {
                Ok(&self.$field)
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