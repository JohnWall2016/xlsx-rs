use super::{ArchiveDeserable, XlsxResult};
use crate::{enum_default, ar_deserable};
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};
use super::zip::Archive;
use super::workbook::{Workbook, Sheet};

pub struct Worksheet {
    
}

impl Worksheet {
    pub fn load_archive(ar: &mut Archive, workbook: &Workbook, sheet: &Sheet) -> XlsxResult<Worksheet> {
        let ws = Worksheet {};
        Ok(ws)
    }
}