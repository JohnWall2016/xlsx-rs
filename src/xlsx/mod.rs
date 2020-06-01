#![allow(dead_code)]

mod zip;
mod workbook;
mod content_types;
mod app_properties;

use serde::Deserialize;
use serde_xml_rs::from_reader;
use std::error::Error;

type Result<T> = std::result::Result<T, Box<dyn Error>>;

fn load_from_zip<'de, T>(ar: &mut zip::Archive, name: &str) -> Result<T>
    where T: Deserialize<'de> {
    let t: T = from_reader(ar.by_name(name)?)?;
    Ok(t)
}

trait LoadArchive: Sized {
    fn load_archive(ar: &mut zip::Archive) -> Result<Self>;
}

#[cfg(test)]
fn test_file() -> String {
    return format!("{}/test_data/table.xlsx", env!("CARGO_MANIFEST_DIR"));
}