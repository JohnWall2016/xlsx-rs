use super::workbook;
use super::worksheet;

use super::{XlsxResult, SharedData};

use super::trie::Map;

pub struct Row {
    book_data: SharedData<workbook::Book>,
    sheet_data: SharedData<worksheet::Sheet>,

    row_data: worksheet::Row,

    cells: Map<Cell>,
}

impl Row {
    pub fn load(
        mut row_data: worksheet::Row,
        sheet_data: SharedData<worksheet::Sheet>,
        book_data: SharedData<workbook::Book>
    ) -> XlsxResult<Row> {
        let mut cells = Map::new();

        for col in row_data.columns.drain(0..) {
            let cell = Cell::load(col)?;
            cells.insert(cell.index(), cell);
        }

        Ok(Row {
            row_data,
            book_data,
            sheet_data,
            cells,
        })
    }

    pub fn index(&self) -> usize {
        self.row_data.reference
    }
}

pub struct Cell {
    column_data: worksheet::Column,
}

impl Cell {
    fn load(
        column_data: worksheet::Column
    ) -> XlsxResult<Cell> {
        Ok(Cell {
            column_data,
        })
    }

    pub fn index(&self) -> usize {
        0 // TODO
    }
}