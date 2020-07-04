#![allow(dead_code)]

use std::io::Read;
use std::error::Error;
use std::fmt::{Debug, Display, Formatter, Result};

use yaserde::{de, ser, YaDeserialize, YaSerialize};

type XlsxErr = Box<dyn Error>;

type XlsxResult<T> = std::result::Result<T, XlsxErr>;

pub struct XlsxError {
    msg: String,
}

impl XlsxError {
    fn new(msg: String) -> XlsxError {
        XlsxError { msg }
    }

    fn error(msg: String) -> XlsxErr {
        XlsxErr::from(
            Self::new(msg)
        )
    }
}

impl Error for XlsxError {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        None
    }

    fn description(&self) -> &str {
        &self.msg
    }

    fn cause(&self) -> Option<&dyn Error> {
        None
    }
}

impl Debug for XlsxError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "xlsx error: {}", self.msg)
    }
}

impl Display for XlsxError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "xlsx error: {}", self.msg)
    }
}

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

use std::cell::{RefCell, Ref, RefMut};
use std::rc::Rc;

#[derive(Debug)]
pub struct SharedData<T>(Rc<RefCell<T>>);

impl<T> SharedData<T> {
    fn new(t: T) -> Self {
        SharedData(Rc::new(RefCell::new(t)))
    }

    fn clone(&self) -> Self {
        SharedData(self.0.clone())
    }

    fn borrow(&self) -> Ref<'_, T> {
        self.0.borrow()
    }

    fn borrow_mut(&self) -> RefMut<'_, T> {
        self.0.borrow_mut()
    }
}

mod map;

mod zip;
mod content_types;
mod app_properties;
mod core_properties;
mod relationships;
mod shared_strings;
mod style_sheet;
mod worksheet;
mod address_converter;
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