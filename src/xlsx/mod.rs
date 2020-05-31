#![allow(dead_code)]

mod workbook;
mod content_types;
mod zip;

#[cfg(test)]
fn test_file() -> String {
    return format!("{}/test_data/table.xlsx", env!("CARGO_MANIFEST_DIR"));
}