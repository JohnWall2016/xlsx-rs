use super::workbook;
use super::worksheet;

use super::{XlsxResult, SharedData};

pub struct Row {
    book: SharedData<workbook::Book>,
    sheet: SharedData<worksheet::Sheet>,

    columns: Vec<Column>,
}

impl Row {
    pub fn load(
        row: &worksheet::Row,
        sheet: SharedData<worksheet::Sheet>,
        book: SharedData<workbook::Book>
    ) -> XlsxResult<Row> {
        let columns = vec![];
        Ok(Row {
            book,
            sheet,
            columns,
        })
    }
}

pub struct Column {

}