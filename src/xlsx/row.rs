use super::workbook;
use super::worksheet;
use super::{XlsxResult, SharedData};
use super::map::IndexMap;
use super::address_converter::CellRef;

pub struct Row {
    book_data: SharedData<workbook::Book>,
    sheet_data: SharedData<worksheet::Sheet>,

    row_data: worksheet::Row,

    pub(crate) cells: IndexMap<Cell>,
}

impl Row {
    pub fn load(
        mut row_data: worksheet::Row,
        sheet_data: SharedData<worksheet::Sheet>,
        book_data: SharedData<workbook::Book>
    ) -> XlsxResult<Row> {
        let mut cells = IndexMap::new();

        for col in row_data.columns.drain(0..) {
            let cell = Cell::load(col, book_data.clone())?;
            cells.put(cell.column_index(), cell);
        }

        Ok(Row {
            row_data,
            book_data,
            sheet_data,
            cells,
        })
    }

    pub fn index(&self) -> usize {
        self.row_data.address_ref
    }

    pub fn cell_at(&self, index: usize) -> &Cell {
        &self.cells[index]
    }
}

pub struct Cell {
    book_data: SharedData<workbook::Book>,
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
        book_data: SharedData<workbook::Book>
    ) -> XlsxResult<Cell> {
        //println!("Cell::load: {:?}", column_data);
        let cell_ref = CellRef::from_address(&column_data.address_ref)?;

        let value = match column_data.typ.as_ref() {
            "s" => {
                if let Ok(idx) = column_data.value.parse::<usize>() {
                    CellValue::String(
                        book_data
                            .borrow()
                            .shared_strings
                            .get_string_by_index(idx).unwrap())
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
            book_data,
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
}