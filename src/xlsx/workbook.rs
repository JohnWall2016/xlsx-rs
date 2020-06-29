use super::content_types::ContentTypes;
use super::app_properties::AppProperties;
use super::core_properties::CoreProperties;
use super::relationships::Relationships;
use super::shared_strings::SharedStrings;
use super::style_sheet::StyleSheet;
use super::worksheet::Worksheet;
use super::zip::Archive;
use super::{XlsxResult, ArchiveDeserable, YaDeserable};

use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

use std::cell::RefCell;
use std::rc::Rc;

pub type SharedBookData = Rc<RefCell<BookData>>;

pub struct Workbook {
    book_data: SharedBookData,

    sheets: Vec<Worksheet>,
}

pub struct BookData {
    content_types: ContentTypes,
    app_properties: AppProperties,
    core_properties: CoreProperties,
    relationships: Relationships,
    shared_strings: SharedStrings,
    style_sheet: StyleSheet,

    book: Book,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "workbook",
    prefix = "",
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    namespace = "r: http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)]
struct Book {
    #[yaserde(rename = "fileVersion")]
    file_version: FileVersion,

    #[yaserde(rename = "workbookPr")]
    property: WorkbookProperty,

    #[yaserde(rename = "workbookProtection")]
    protection: Option<WorkbookProtection>,

    #[yaserde(rename = "bookViews")]
    book_views: BookViews,

    sheets: Sheets,

    #[yaserde(rename = "definedNames")]
    defined_names: Option<DefinedNames>,

    #[yaserde(rename = "calcPr")]
    calc_property: CalcProperty,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct FileVersion {
    #[yaserde(attribute, rename = "appName")]
    app_name: String,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct WorkbookProperty {
    #[yaserde(attribute, rename = "showObjects")]
    show_objects: Option<String>,

    #[yaserde(attribute)]
    date1904: Option<String>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct WorkbookProtection {}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct BookViews {
    #[yaserde(rename = "workbookView")]
    items: Vec<WorkbookView>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(rename = "workbookView")]
struct WorkbookView {
    #[yaserde(attribute, rename = "showHorizontalScroll")]
    show_horizontal_scroll: Option<String>,

    #[yaserde(attribute, rename = "showVerticalScroll")]
    show_vertical_scroll: Option<String>,

    #[yaserde(attribute, rename = "showSheetTabs")]
    show_sheet_tabs: Option<String>,

    #[yaserde(attribute, rename = "tabRatio")]
    tab_ratio: Option<String>,

    #[yaserde(attribute, rename = "windowHeight")]
    window_height: Option<String>,

    #[yaserde(attribute, rename = "windowWidth")]
    window_width: Option<String>,

    #[yaserde(attribute, rename = "xWindow")]
    x_window: Option<String>,

    #[yaserde(attribute, rename = "yWindow")]
    y_window: Option<String>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct Sheets {
    #[yaserde(rename = "sheet")]
    items: Vec<Sheet>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    rename = "sheet"
    prefix = "",
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    namespace = "r: http://schemas.openxmlformats.org/officeDocument/2006/relationships",
)]
pub struct Sheet {
    #[yaserde(attribute)]
    name: String,

    #[yaserde(attribute, rename = "sheetId")]
    pub(crate) sheet_id: u32,

    #[yaserde(attribute, prefix = "r")]
    id: String,

    #[yaserde(attribute)]
    state: Option<String>,
}

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct DefinedNames { }

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct CalcProperty {
    #[yaserde(attribute, rename = "iterateCount")]
    iterate_count: Option<String>,

    #[yaserde(attribute, rename = "refMode")]
    ref_mode: Option<String>,

    #[yaserde(attribute, rename = "iterateDelta")]
    iterate_delta: Option<String>,

    #[yaserde(attribute, rename = "calcId")]
    calc_id: Option<String>,
}

impl ArchiveDeserable for Workbook {
    fn path() -> &'static str {
        "xl/workbook.xml"
    }

    fn from_archive(ar: &mut Archive) -> XlsxResult<Workbook> {
        let mut wb = BookData {
            content_types:  ContentTypes::from_archive(ar)?,
            app_properties: AppProperties::from_archive(ar)?,
            core_properties: CoreProperties::from_archive(ar)?,
            relationships: Relationships::from_archive(ar)?,
            shared_strings: SharedStrings::from_archive(ar)?,
            style_sheet: StyleSheet::from_archive(ar)?,
            book: Book::from_reader(Self::archive_reader(ar)?)?,
        };

        if wb.relationships.find_by_type("sharedStrings").is_none() {
            wb.relationships.add("sharedStrings", "sharedStrings.xml");
        }

        if wb.content_types.find_by_part_name("/xl/sharedStrings.xml").is_none() {
            wb.content_types.add(
                "/xl/sharedStrings.xml", 
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
            );
        }

        let book_data = Rc::new(RefCell::new(wb));
        let mut sheets = vec![];

        for sheet in &book_data.borrow().book.sheets.items {
            sheets.push(
                Worksheet::load_archive(ar, book_data.clone(), sheet.sheet_id)?
            );
        }

        Ok(Workbook {
            book_data,
            sheets,
        })
    }

    fn to_string(&self) -> XlsxResult<String> {
        Ok(self.book_data.borrow().book.to_string()?)
    }
}

#[test]
fn test_load_ar() -> super::XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", Workbook::archive_string(&mut ar)?);

    let wb = Workbook::from_archive(&mut ar)?;
    println!("{:?}\n", wb.book_data.borrow().book);

    println!("{}\n", wb.to_string()?);

    Ok(())
}