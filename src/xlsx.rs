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
pub struct Sheet {}

#[derive(Debug)]
pub struct Row {}

#[derive(Debug)]
pub struct Col {}

use xml::sheet::Worksheet;
use result::{XlsxResult, Error};
use refer;

impl Sheet {
    pub fn from_xml(
        worksheet: Worksheet,
        strs: &refer::Strings,
        clrs: &refer::Colors,
        nfts: &refer::NumFmts,
    ) -> XlsxResult<Self> {
        let sheet = Sheet {};
        let (min_col, min_row, max_col, max_row) = if worksheet.dimension.refer != "" {
            Self::get_boundary_from_dimenref(&worksheet.dimension.refer)?
        } else {
            Self::calc_boundary_from_worksheet(&worksheet)?
        };

        println!("{:?}", (min_col, min_row, max_col, max_row));
        Ok(sheet)
    }

    fn get_boundary_from_dimenref(refer: &String) -> XlsxResult<(usize, usize, usize, usize)> {
        let parts: Vec<&str> = refer.split(":").collect();
        if parts.len() != 2 {
            return Error::xlsx("sheet dimension format error");
        }

        let (minx, miny) = match Self::get_coords_from_cellstr(parts[0]) {
            Some((x, y)) => (x, y),
            None => return Error::xlsx("sheet dimension format error"),
        };

        let (maxx, maxy) = match Self::get_coords_from_cellstr(parts[1]) {
            Some((x, y)) => (x, y),
            None => return Error::xlsx("sheet dimension format error"),
        };

        Ok((minx, miny, maxx, maxy))
    }

    fn get_coords_from_cellstr(cs: &str) -> Option<(usize, usize)> {
        let x = match Self::letters_to_numbers(cs.trim_right_matches(char::is_numeric)) {
            Some(n) => n,
            None => return None,
        };

        let y = match cs.trim_left_matches(|c| c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z')
            .parse::<usize>() {
            Ok(n) => n,
            Err(_) => return None,
        };

        Some((x - 1, y - 1))
    }

    fn letters_to_numbers(lt: &str) -> Option<usize> {
        let mut num: usize = 0;
        for b in lt.bytes() {
            if b >= b'a' && b <= b'z' {
                num = num * 26 + usize::from(b - b'a') + 1;
            } else if b >= b'A' && b <= b'Z' {
                num = num * 26 + usize::from(b - b'A') + 1;
            } else {
                return None;
            }
        }
        Some(num)
    }

    fn calc_boundary_from_worksheet(
        worksheet: &Worksheet,
    ) -> XlsxResult<(usize, usize, usize, usize)> {
        let (mut minx, mut miny, mut maxx, mut maxy) =
            (usize::max_value(), usize::max_value(), 0, 0);
        for row in worksheet.sheetData.items() {
            for col in row.items() {
                let (x, y) = match Self::get_coords_from_cellstr(&col.r) {
                    Some((x, y)) => (x, y),
                    None => return Error::xlsx("sheet data format error"),
                };
                if x < minx {
                    minx = x;
                }
                if x > maxx {
                    maxx = x;
                }
                if y < miny {
                    miny = y;
                }
                if y > maxx {
                    maxy = y;
                }
            }
        }
        if minx == usize::max_value() || miny == usize::max_value() {
            return Error::xlsx("cannot get boundary from worksheet");
        }
        Ok((minx, miny, maxx, maxy))
    }
}
