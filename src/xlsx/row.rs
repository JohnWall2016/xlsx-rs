use std::ops::{Index, IndexMut};

use crate::xlsx::base::XlsxError;

use super::address_converter::{CellRef, ToNumber};
use super::base::{SharedData, XlsxResult};
use super::map::IndexMap;
use super::workbook;
use super::worksheet;

pub struct Row {
    book_shared_data: SharedData<workbook::Book>,
    sheet_shared_data: SharedData<worksheet::Sheet>,

    row_data: worksheet::Row,

    pub(crate) cells: IndexMap<Cell>,
}

impl Row {
    pub fn load(
        mut row_data: worksheet::Row,
        sheet_shared_data: SharedData<worksheet::Sheet>,
        book_shared_data: SharedData<workbook::Book>,
    ) -> XlsxResult<Row> {
        let mut cells = IndexMap::new();

        for col in row_data.columns.drain(0..) {
            let cell = Cell::load(col, book_shared_data.clone())?;
            cells.put(cell.column_index(), cell);
        }

        Ok(Row {
            row_data,
            book_shared_data,
            sheet_shared_data,
            cells,
        })
    }

    pub fn index(&self) -> usize {
        self.row_data.address_ref
    }

    pub fn get_cell<P: ToNumber>(&self, index: P) -> XlsxResult<&Cell> {
        let index = index.to_number();
        match self.cells.get(index) {
            Some(v) => Ok(v),
            None => Err(XlsxError::error(format!("out of index: {}", index))),
        }
    }

    pub fn get_cell_mut<P: ToNumber>(&mut self, index: P) -> XlsxResult<&mut Cell> {
        let index = index.to_number();
        match self.cells.get_mut(index) {
            Some(v) => Ok(v),
            None => Err(XlsxError::error(format!("out of index: {}", index))),
        }
    }
}

impl<P: ToNumber> Index<P> for Row {
    type Output = Cell;

    #[inline]
    fn index(&self, col: P) -> &Cell {
        &self.cells[col.to_number()]
    }
}

impl<P: ToNumber> IndexMut<P> for Row {
    #[inline]
    fn index_mut(&mut self, col: P) -> &mut Cell {
        &mut self.cells[col.to_number()]
    }
}

pub struct Cell {
    book_shared_data: SharedData<workbook::Book>,
    column_data: worksheet::Column,
    cell_ref: CellRef,
    value: CellValue,
}

#[derive(Debug)]
pub enum CellValue {
    String(String),
    Bool(bool),
    Error(FormulaError),
    Number(f64),
    None,
}

impl CellValue {
    pub fn as_str(&self) -> Option<&str> {
        if let CellValue::String(s) = self {
            Some(s)
        } else {
            None
        }
    }

    pub fn as_bool(&self) -> Option<&bool> {
        if let CellValue::Bool(b) = self {
            Some(b)
        } else {
            None
        }
    }

    pub fn as_num(&self) -> Option<&f64> {
        if let CellValue::Number(f) = self {
            Some(f)
        } else {
            None
        }
    }
}

#[derive(Debug)]
pub struct FormulaError {
    error: String,
}

impl FormulaError {
    fn new(error: String) -> FormulaError {
        FormulaError { error }
    }

    fn error(&self) -> &str {
        &self.error
    }
}

impl Cell {
    fn load(
        column_data: worksheet::Column,
        book_shared_data: SharedData<workbook::Book>,
    ) -> XlsxResult<Cell> {
        //println!("Cell::load: {:?}", column_data);
        let cell_ref = CellRef::from_address(&column_data.address_ref)?;

        let value = match column_data.typ.as_ref() {
            "s" => {
                if let Ok(idx) = column_data.value.parse::<usize>() {
                    CellValue::String(
                        book_shared_data
                            .borrow()
                            .shared_strings
                            .get_string_by_index(idx)
                            .unwrap(),
                    )
                } else {
                    CellValue::String(String::from(""))
                }
            }
            "str" => CellValue::String(column_data.value.clone()),
            "inlineStr" => unimplemented!(),
            "b" => CellValue::Bool(column_data.value == "1"),
            "e" => CellValue::Error(FormulaError::new(column_data.value.clone())),
            _ => {
                let num = column_data.value.parse::<f64>();
                if let Ok(f) = num {
                    CellValue::Number(f)
                } else {
                    CellValue::None
                }
            }
        };

        Ok(Cell {
            book_shared_data,
            column_data,
            cell_ref,
            value,
        })
    }

    pub fn column_index(&self) -> usize {
        self.cell_ref.column()
    }

    pub fn row_index(&self) -> usize {
        self.cell_ref.row()
    }

    pub fn value(&self) -> &CellValue {
        &self.value
    }

    pub fn set_value_string(&mut self, value: String) {
        self.set_value(CellValue::String(value))
    }

    pub fn set_value_bool(&mut self, value: bool) {
        self.set_value(CellValue::Bool(value))
    }

    pub fn set_value_number(&mut self, value: f64) {
        self.set_value(CellValue::Number(value))
    }

    pub fn set_value(&mut self, value: CellValue) {
        match &value {
            CellValue::String(s) => {
                self.column_data.typ = "s".to_string();
                let index = self
                    .book_shared_data
                    .borrow_mut()
                    .shared_strings
                    .get_index_for_string(&s);
                self.column_data.value = format!("{}", index);
            }
            CellValue::Bool(b) => {
                self.column_data.typ = "b".to_string();
                self.column_data.value = if *b { "1".to_string() } else { "0".to_string() }
            }
            CellValue::Number(f) => {
                self.column_data.typ = "".to_string();
                self.column_data.value = f.to_string();
            }
            _ => return,
        }
        self.value = value;
    }
}
