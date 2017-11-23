use std::collections::BTreeMap as Map;

#[derive(Debug)]
pub struct WorkBook {
    pub date1904: bool,
    sheets: Vec<Sheet>,
    names_map: Map<String, usize>,
}

impl WorkBook {
    pub fn new() -> Self {
        WorkBook {
            date1904: false,
            sheets: Vec::new(),
            names_map: Map::new(),
        }
    }

    pub fn insert(&mut self, name: &String, sheet: Sheet) {
        self.sheets.push(sheet);
        self.names_map.insert(name.clone(), self.sheets.len() - 1);
    }
}

#[derive(Debug)]
pub struct Sheet {
    
}

use xml::sheet::Worksheet;
use result::{XlsxResult, Error};
use refer;

impl Sheet {
    pub fn from_xml(worksheet: Worksheet,
                    strs: &refer::Strings,
                    clrs: &refer::Colors,
                    nfts: &refer::NumFmts
    ) -> XlsxResult<Self> {
        let sheet = Sheet{};
        let bdy = Self::get_boundary_from_dimenref(&worksheet.dimension.refer);
        println!("{:?}", bdy);
        Ok(sheet)
    }

    fn get_boundary_from_dimenref(refer: &String) -> XlsxResult<(usize, usize, usize, usize)> {
        let parts: Vec<&str> = refer.split(":").collect();
        if parts.len() != 2 {
            return Error::xlsx("sheet dimension format error");
        }
        let top: usize = parts[0].trim_left_matches(
            |c| c >= 'A' && c <='Z' || c >= 'a' && c <= 'z'
        ).parse().or(Error::xlsx("sheet dimension format error"))?;
        let bottom: usize = parts[1].trim_left_matches(
            |c| c >= 'A' && c <='Z' || c >= 'a' && c <= 'z'
        ).parse().or(Error::xlsx("sheet dimension format error"))?;

        Ok((0, top, 0, bottom))
    }
}
