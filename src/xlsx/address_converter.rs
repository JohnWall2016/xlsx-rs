use regex::Regex;

use std::fmt::Write;

use super::{XlsxResult, XlsxError};

const CELL_RE_STR: &str = r"(\$?)([A-Z]+)(\$?)(\d+)";

lazy_static! {
    static ref CELL_REGEX: Regex = Regex::new(CELL_RE_STR).unwrap();
    static ref RANGE_REGEX: Regex = Regex::new(
        &format!("({0}):({0})", CELL_RE_STR)
    ).unwrap();
}

#[derive(Debug)]
pub struct CellRef {
    row: usize,
    column: usize,
    column_name: String,
    row_anchored: bool,
    column_anchored: bool,
}

impl CellRef {
    pub fn new(row: usize, column: usize, row_anchored: bool, column_anchored: bool) -> CellRef {
        CellRef {
            row,
            column,
            column_name: column_number_to_name(column),
            row_anchored,
            column_anchored,
        }
    }

    pub fn from_address(address: &str) -> XlsxResult<CellRef> {
        let result = CELL_REGEX.captures(address);

        if let Some(caps) = result {
            let column_anchored = !caps[1].is_empty();
            let column_name = caps[2].to_string();
            let row_anchored = !caps[3].is_empty();
            let row = caps[4].parse::<usize>()?;
            let column = column_name_to_number(&column_name);
            Ok(CellRef{
                row,
                column,
                column_name,
                row_anchored,
                column_anchored,
            })
        } else {
            Err(
                XlsxError::error(
                format!("invalid address: \"{}\"", address)
                )
            )
        }
    }

    pub fn to_address(&self) -> String {
        let mut s = String::new();
        if self.column_anchored {
            write!(s, "$").unwrap();
        }
        write!(s, "{}", self.column_name).unwrap();
        if self.row_anchored {
            write!(s, "$").unwrap();
        }
        write!(s, "{}", self.row).unwrap();
        s
    }

    pub fn row(&self) -> usize {
        self.row
    }

    pub fn column(&self) -> usize {
        self.column
    }

    pub fn column_name(&self) -> &str {
        &self.column_name
    }

    pub fn row_anchored(&self) -> bool {
        self.row_anchored
    }

    pub fn column_anchored(&self) -> bool {
        self.column_anchored
    }
}

#[derive(Debug)]
pub struct RangeRef {
    start: CellRef,
    end: CellRef,
}

impl RangeRef {
    pub fn new(start: CellRef, end: CellRef) -> RangeRef {
        RangeRef {
            start,
            end,
        }
    }

    pub fn from_address(address: &str) -> XlsxResult<RangeRef> {
        let result = RANGE_REGEX.captures(address);

        if let Some(caps) = result {
            Ok(RangeRef {
                start: CellRef::from_address(&caps[1])?,
                end: CellRef::from_address(&caps[6])?,
            })
        } else {
            Err(
                XlsxError::error(
                format!("invalid address: \"{}\"", address)
                )
            )
        }
    }

    pub fn to_address(&self) -> String {
        format!("{}:{}", self.start.to_address(), self.end.to_address())
    }
}

fn column_name_to_number(name: &str) -> usize {
    let name = name.to_uppercase();
    let mut index = 0 as usize;
    for ch in name.chars() {
        index *= 26;
        index += ch as usize - 'A' as usize + 1;
    }
    index
}

fn column_number_to_name(index: usize) -> String {
    let mut dividend = index;
    let mut name = String::new();

    while dividend > 0 {
        let remainder = (dividend - 1) % 26;
        name = format!("{}", (65 + remainder) as u8 as char) + &name;
        dividend = (dividend - remainder) / 26;
    }

    name
}


#[test]
fn test_regex_captures() {
    let caps = CELL_REGEX.captures("$A$1");
    println!("{:?}", caps);
    let caps = CELL_REGEX.captures("A1");
    println!("{:?}", caps);
    let caps = CELL_REGEX.captures("A");
    println!("{:?}", caps);

    let cell = CellRef::from_address("$A$1").unwrap();
    println!("{:?}, {}", cell, cell.to_address());
    let cell = CellRef::from_address("B2").unwrap();
    println!("{:?}, {}", cell, cell.to_address());
    println!("{:?}", CellRef::from_address("C"));

    println!("{} {}", column_number_to_name(1), column_name_to_number("A"));
    println!("{} {}", column_number_to_name(28), column_name_to_number("AB"));

    let range = RangeRef::from_address("A1:AB22").unwrap();
    println!("{:?}, {}", range, range.to_address());
}
