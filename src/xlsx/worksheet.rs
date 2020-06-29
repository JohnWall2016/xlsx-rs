use super::{YaDeserable, XlsxResult};
use crate::enum_default;
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};
use super::zip::{Archive, ReadAll};
use super::workbook;

pub struct Worksheet {
    sheet: Sheet,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "worksheet",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
pub struct Sheet {
    #[yaserde(rename = "sheetPr")]
    sheet_properties: SheetProperties,


}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "sheetPr")]
struct SheetProperties {
    #[yaserde(attribute, rename = "filterMode")]
    filter_mode: String,

    #[yaserde(rename = "pageSetUpPr")]
    page_setup_properties: PageSetupProperties,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "pageSetUpPr")]
struct PageSetupProperties {
    #[yaserde(attribute, rename = "fitToPage")]
    fit_to_page: String,
}

impl Worksheet {
    pub fn load_archive(ar: &mut Archive, workbook: &workbook::Workbook, sheet: &workbook::Sheet) -> XlsxResult<Worksheet> {
        let path = format!("xl/worksheets/sheet{}.xml", sheet.sheet_id);

        println!("sheet: {}\n", path);

        println!("{}\n", ar.by_name(&path)?.read_all_to_string()?);
        
        let sheet = Sheet::from_reader(ar.by_name(&path)?)?;

        println!("{:?}\n", sheet);

        println!("{}\n", sheet.to_string()?);
        
        Ok(Worksheet {
            sheet,
        })
    }
}