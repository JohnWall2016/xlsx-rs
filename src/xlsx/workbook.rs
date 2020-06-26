use super::content_types::ContentTypes;
use super::app_properties::AppProperties;
use super::core_properties::CoreProperties;
use super::relationships::Relationships;
use super::shared_strings::SharedStrings;
use super::style_sheet::StyleSheet;
use super::zip::Archive;
use super::{XlsxResult, ArchiveDeserable};

use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

use yaserde::de::from_reader;
use yaserde::ser::to_string;

pub struct Workbook {
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
struct Sheet {
    #[yaserde(attribute)]
    name: String,

    #[yaserde(attribute, rename = "sheetId")]
    sheet_id: String,

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

impl Workbook {
    fn path() -> &'static str {
        "xl/workbook.xml"
    }

    fn archive_str(ar: &mut Archive) -> XlsxResult<String> {
        use super::zip::ReadAll;
        Ok(ar.by_name(Self::path())?.read_all_to_string()?)
    }

    fn load_archive(ar: &mut Archive) -> XlsxResult<Workbook> {
        let book: Book = from_reader(ar.by_name(Self::path())?)?;

        let mut wb = Workbook {
            content_types:  ContentTypes::load_archive(ar)?,
            app_properties: AppProperties::load_archive(ar)?,
            core_properties: CoreProperties::load_archive(ar)?,
            relationships: Relationships::load_archive(ar)?,
            shared_strings: SharedStrings::load_archive(ar)?,
            style_sheet: StyleSheet::load_archive(ar)?,
            book,
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

        Ok(wb)
    }

    fn to_string(&self) -> XlsxResult<String> {
        Ok(to_string(&self.book)?)
    }
}

#[test]
fn test_load_ar() -> super::XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", Workbook::archive_str(&mut ar)?);

    let wb = Workbook::load_archive(&mut ar)?;
    println!("{:?}\n", wb.book);

    println!("{}\n", wb.to_string()?);

    Ok(())
}