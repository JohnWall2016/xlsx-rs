use super::*;

pub fn test_file() -> String {
    return format!("{}/test_data/table.xlsx", env!("CARGO_MANIFEST_DIR"));
}

pub fn test_archive() -> base::XlsxResult<zip::Archive> {
    Ok(zip::Archive::new(test_file())?)
}
