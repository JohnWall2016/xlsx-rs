use super::workbook;
use super::worksheet;

use super::XlsxResult;

pub struct Row {
    book: workbook::SharedBookData,
    sheet: worksheet::SharedSheetData,

    columns: Vec<Column>,
}

impl Row {
    pub fn load(row: &worksheet::Row, sheet: worksheet::SharedSheetData, book: workbook::SharedBookData) -> XlsxResult<Row> {
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