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

    pub fn insert(&mut self, name: &str, sheet: Sheet) {
        self.sheets.push(sheet);
        self.names_map.insert(name.into(), self.sheets.len() - 1);
    }
}

#[derive(Debug)]
pub struct Sheet {
    cols: Vec<Col>,

    max_col: usize,
    max_row: usize,
}

use xml::sheet::Worksheet as XmlWorksheet;
use xml::sheet::Cols as XmlSheetCols;
use result::{XlsxResult, Error};
//use refer;
use file;

use std::cmp;

impl Sheet {
    pub fn from_xml(worksheet: XmlWorksheet, file: &file::File) -> XlsxResult<Self> {
        let mut sheet = Sheet {
            cols: Vec::new(),
            max_col: 0,
            max_row: 0,
        };
        let (min_col, min_row, max_col, max_row) = if worksheet.dimension.refer != "" {
            Self::get_boundary_from_dimenref(&worksheet.dimension.refer)?
        } else {
            Self::calc_boundary_from_worksheet(&worksheet)?
        };
        //println!("{:?}", (min_col, min_row, max_col, max_row));
        sheet.max_col = max_col;
        sheet.max_row = max_row;

        for _ in 0..max_col + 1 {
            sheet.cols.push(Col::new());
        }

        if worksheet.cols.is_some() {
            sheet.update_cols_from_worksheet(
                &worksheet.cols.unwrap(),
                file,
            )?;
        }

        Ok(sheet)
    }

    fn get_boundary_from_dimenref(refer: &str) -> XlsxResult<(usize, usize, usize, usize)> {
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
        worksheet: &XmlWorksheet,
    ) -> XlsxResult<(usize, usize, usize, usize)> {
        let (mut minx, mut miny, mut maxx, mut maxy) =
            (usize::max_value(), usize::max_value(), 0, 0);
        for row in &worksheet.sheetData.items {
            for col in &row.items {
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

    fn update_cols_from_worksheet(
        &mut self,
        cols: &XmlSheetCols,
        file: &file::File,
    ) -> XlsxResult<()> {
        for col in &cols.items {
            let min = col.min.parse::<usize>()? - 1;
            let max = col.max.parse::<usize>()? - 1;
            for i in min..cmp::min(max, self.max_col) + 1 {
                self.cols[i].min = min;
                self.cols[i].max = max;
                self.cols[i].hidden = col.hidden.parse()?;
                self.cols[i].collapsed = col.collapsed.parse()?;
                self.cols[i].width = col.width.parse()?;
                self.cols[i].outline_level = Self::string_option_parse(&col.outlineLevel)?;
                self.cols[i].num_fmt = file.get_num_fmt(col.style.parse()?);
            }
        }
        Ok(())
    }

    fn string_option_parse<T: ::std::str::FromStr>(
        opt: &Option<String>,
    ) -> Result<Option<T>, T::Err> {
        match *opt {
            Some(ref s) => {
                match s.parse() {
                    Ok(t) => Ok(Some(t)),
                    Err(err) => Err(err),
                }
            }
            None => Ok(None),
        }
    }
}

#[derive(Debug)]
pub struct Row {}

#[derive(Debug)]
pub struct Col {
    min: usize,
    max: usize,
    hidden: bool,
    width: f64,
    collapsed: bool,
    outline_level: Option<u8>,
    num_fmt: Option<String>,
    //style:
}

impl Col {
    fn new() -> Self {
        Col {
            min: 0,
            max: 0,
            hidden: false,
            width: 0f64,
            collapsed: false,
            outline_level: None,
            num_fmt: None,
        }
    }
}
