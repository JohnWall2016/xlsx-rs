use super::{YaDeserable, XlsxResult, SharedData};
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};
use super::zip::{Archive, ReadAll};
use super::workbook;
use super::row;

pub struct Worksheet {
    book_data: SharedData<workbook::Book>,
    sheet_data: SharedData<Sheet>,

    rows: Vec<row::Row>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "worksheet",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
pub struct Sheet {
    #[yaserde(rename = "sheetPr")]
    sheet_properties: SheetProperties,

    #[yaserde(rename = "dimension")]
    dimension: Dimension,

    #[yaserde(rename = "sheetViews")]
    sheet_views: SheetViews,

    #[yaserde(rename = "sheetFormatPr")]
    sheet_format_properties: SheetFormatProperties,

    cols: Cols,

    #[yaserde(rename = "sheetData")]
    sheet_data: Option<SheetData>,

    #[yaserde(rename = "printOptions")]
    print_options: PrintOptions,

    #[yaserde(rename = "pageMargins")]
    page_margins: PageMargins,

    #[yaserde(rename = "pageSetup")]
    page_setup: PageSetup,

    #[yaserde(rename = "headerFooter")]
    header_footer: HeaderFooter,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "sheetPr")]
struct SheetProperties {
    #[yaserde(attribute, rename = "filterMode")]
    filter_mode: String,

    #[yaserde(rename = "pageSetUpPr")]
    page_setup_properties: PageSetupProperties,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "pageSetUpPr")]
struct PageSetupProperties {
    #[yaserde(attribute, rename = "fitToPage")]
    fit_to_page: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "dimension")]
struct Dimension {
    #[yaserde(attribute, rename = "ref")]
    reference: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "sheetViews")]
struct SheetViews {
    #[yaserde(rename = "sheetView")]
    items: Vec<SheetView>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "sheetView")]
struct SheetView {
    #[yaserde(attribute, rename = "windowProtection")]
    window_protection: String,

    #[yaserde(attribute, rename = "showFormulas")]
    show_formulas: String,

    #[yaserde(attribute, rename = "showGridLines")]
    show_grid_lines: String,

    #[yaserde(attribute, rename = "showRowColHeaders")]
    show_row_col_headers: String,

    #[yaserde(attribute, rename = "showZeros")]
    show_zeros: String,

    #[yaserde(attribute, rename = "rightToLeft")]
    right_to_left: String,

    #[yaserde(attribute, rename = "tabSelected")]
    tab_selected: String,

    #[yaserde(attribute, rename = "showOutlineSymbols")]
    show_outline_symbols: String,

    #[yaserde(attribute, rename = "view")]
    view: String,

    #[yaserde(attribute, rename = "topLeftCell")]
    top_left_cell: String,

    #[yaserde(attribute, rename = "zoomScale")]
    zoom_scale: String,

    #[yaserde(attribute, rename = "zoomScaleNormal")]
    zoom_scale_normal: String,

    #[yaserde(attribute, rename = "zoomScalePageLayoutView")]
    zoom_scale_page_layout_view: String,

    #[yaserde(attribute, rename = "workbookViewId")]
    workbook_view_id: String,

    selection: Selection,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "selection")]
struct Selection {
    #[yaserde(attribute, rename = "pane")]
    top_left: String,

    #[yaserde(attribute, rename = "activeCell")]
    active_cell: String,

    #[yaserde(attribute, rename = "activeCellId")]
    active_cell_id: String,

    #[yaserde(attribute, rename = "sqref")]
    sqref: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "sheetFormatPr")]
struct SheetFormatProperties {
    #[yaserde(attribute, rename = "defaultColWidth")]
    default_col_width: String,

    #[yaserde(attribute, rename = "defaultRowHeight")]
    default_row_height: String,

    #[yaserde(attribute, rename = "customHeight")]
    custom_height: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "cols")]
struct Cols {
    #[yaserde(rename = "col")]
    items: Vec<Col>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "col")]
struct Col {
    #[yaserde(attribute, rename = "collapsed")]
    collapsed: String,

    #[yaserde(attribute, rename = "hidden")]
    hidden: String,

    #[yaserde(attribute, rename = "max")]
    max: String,

    #[yaserde(attribute, rename = "min")]
    min: String,

    #[yaserde(attribute, rename = "style")]
    style: String,

    #[yaserde(attribute, rename = "width")]
    width: String,

    #[yaserde(attribute, rename = "customWidth")]
    custom_width: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "sheetData")]
struct SheetData {
    #[yaserde(rename = "row")]
    items: Vec<Row>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "row")]
pub struct Row {
    #[yaserde(attribute, rename = "r")]
    pub reference: u32,

    #[yaserde(attribute, rename = "ht")]
    pub height: String,

    #[yaserde(attribute, rename = "customHeight")]
    pub custom_height: String,

    #[yaserde(rename = "c")]
    pub columns: Vec<Column>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    rename = "c",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
pub struct Column {
    #[yaserde(attribute, rename = "r")]
    pub reference: String,

    #[yaserde(attribute, rename = "s")]
    pub style: String,

    #[yaserde(attribute, rename = "t")]
    pub typ: String,

    #[yaserde(rename = "v")]
    pub value: String
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "printOptions")]
struct PrintOptions {
    #[yaserde(attribute, rename = "horizontalCentered")]
    horizontal_centered: String,

    #[yaserde(attribute, rename = "verticalCentered")]
    vertical_centered: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "pageMargins")]
struct PageMargins {
    #[yaserde(attribute, rename = "left")]
    left: String,

    #[yaserde(attribute, rename = "right")]
    right: String,
    
    #[yaserde(attribute, rename = "top")]
    top: String,

    #[yaserde(attribute, rename = "bottom")]
    bottom: String,
    
    #[yaserde(attribute, rename = "header")]
    header: String,

    #[yaserde(attribute, rename = "footer")]
    footer: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "pageSetup")]
struct PageSetup {
    #[yaserde(attribute, rename = "paperSize")]
    paper_size: String,

    #[yaserde(attribute, rename = "orientation")]
    orientation: String,
    
    #[yaserde(attribute, rename = "horizontalDpi")]
    horizontal_dpi: String,

    #[yaserde(attribute, rename = "verticalDpi")]
    vertical_dpi: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "headerFooter")]
struct HeaderFooter {
    #[yaserde(attribute, rename = "differentFirst")]
    different_first: String,

    #[yaserde(attribute, rename = "differentOddEven")]
    different_odd_even: String,
}

impl Worksheet {
    pub fn load_archive(
        ar: &mut Archive,
        book_data: SharedData<workbook::Book>,
        sheet_id: u32
    ) -> XlsxResult<Worksheet> {
        let path = format!("xl/worksheets/sheet{}.xml", sheet_id);

        println!("sheet: {}\n", path);

        println!("{}\n", ar.by_name(&path)?.read_all_to_string()?);
        
        let sheet = Sheet::from_reader(ar.by_name(&path)?)?;

        println!("{:?}\n", sheet);

        println!("{}\n", sheet.to_string()?);

        let sheet_data = SharedData::new(sheet);

        let mut rows = vec![];
        
        if let Some(data) = sheet_data.borrow_mut().sheet_data.take() {
            for row_data in data.items {
                rows.push(
                    row::Row::load(row_data, sheet_data.clone(), book_data.clone())?
                )
            }
        }
        
        Ok(Worksheet {
            book_data,
            sheet_data,
            rows,
        })
    }
}