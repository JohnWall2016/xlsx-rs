#![allow(dead_code)]

mod zip;
mod workbook;
mod content_types;
//mod app_properties;

use yaserde::de::{from_str, from_reader};
use yaserde::ser::{to_string, to_string_content};
use yaserde::{YaDeserialize, YaSerialize};

use std::error::Error;

type XlsXResult<T> = std::result::Result<T, Box<dyn Error>>;

fn load_from_zip<T>(ar: &mut zip::Archive, name: &str) -> XlsXResult<T>
    where T: YaDeserialize {
    let t: T = from_reader(ar.by_name(name)?)?;
    Ok(t)
}

trait LoadArchive: Sized {
    fn load_archive(ar: &mut zip::Archive) -> XlsXResult<Self>;
}

#[cfg(test)]
fn test_file() -> String {
    return format!("{}/test_data/table.xlsx", env!("CARGO_MANIFEST_DIR"));
}