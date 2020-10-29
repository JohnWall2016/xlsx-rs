use super::content_types::ContentTypes;
use super::app_properties::AppProperties;
use super::core_properties::CoreProperties;
use super::relationships::Relationships;
use super::shared_strings::SharedStrings;
use super::style_sheet::StyleSheet;
use super::worksheet::Worksheet;
use super::zip::Archive;
use super::base::{XlsxResult, ArchiveDeserable, YaDeserable, SharedData};
use super::map::IndexMap;

use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

pub struct Workbook {
    book_shared_data: SharedData<Book>,

    sheets: IndexMap<Worksheet>,
}

pub struct Book {
    content_types: ContentTypes,
    app_properties: AppProperties,
    core_properties: CoreProperties,
    relationships: Relationships,
    pub(crate) shared_strings: SharedStrings,
    style_sheet: StyleSheet,

    book: WBook,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "workbook",
    prefix = "",
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    namespace = "r: http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)]
struct WBook {
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
    pub(crate) sheet_id: usize,

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
        let mut book = Book {
            content_types:  ContentTypes::from_archive(ar)?,
            app_properties: AppProperties::from_archive(ar)?,
            core_properties: CoreProperties::from_archive(ar)?,
            relationships: Relationships::from_archive(ar)?,
            shared_strings: SharedStrings::from_archive(ar)?,
            style_sheet: StyleSheet::from_archive(ar)?,
            book: WBook::from_reader(Self::archive_reader(ar)?)?,
        };

        if book.relationships.find_by_type("sharedStrings").is_none() {
            book.relationships.add("sharedStrings", "sharedStrings.xml");
        }

        if book.content_types.find_by_part_name("/xl/sharedStrings.xml").is_none() {
            book.content_types.add(
                "/xl/sharedStrings.xml", 
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
            );
        }

        let book_shared_data = SharedData::new(book);
        let mut sheets = IndexMap::new();

        for (index, sheet) in (&book_shared_data.borrow().book.sheets).items.iter().enumerate() {
            sheets.put(
                index,
                Worksheet::load_archive(ar, book_shared_data.clone(), sheet.sheet_id)?
            );
        }

        Ok(Workbook {
            book_shared_data,
            sheets,
        })
    }

    fn to_string(&self) -> XlsxResult<String> {
        Ok(self.book_shared_data.borrow().book.to_string()?)
    }
}

impl Workbook {
    pub fn sheet_at(&self, index: usize) -> &Worksheet {
        &self.sheets[index]
    }

    pub fn sheet_mut_at(&mut self, index: usize) -> &mut Worksheet {
        &mut self.sheets[index]
    }
}

#[test]
fn test_load_ar() -> XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    //println!("{}\n", Workbook::archive_string(&mut ar)?);

    let mut wb = Workbook::from_archive(&mut ar)?;
    //println!("{:?}\n", wb.book_data.borrow().book);

    //println!("{}\n", wb.to_string()?);

    println!("{:?}", wb.sheet_at(0).row_at(1).cell_at(1).value().as_str().unwrap());
    println!("{:?}", wb.sheet_at(0).cell("C5")?.value().as_str().unwrap());

    wb.sheet_mut_at(0).cell_mut("C5")?.set_value_string("abc中国人".to_string());
    println!("{:?}", wb.sheet_at(0).cell("C5")?.value().as_str().unwrap());
    Ok(())
}