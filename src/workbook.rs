#[derive(Debug, Deserialize)]
struct Workbook {
    fileVersion: FileVersion,
    workbookPr: WorkbookPr,
    workbookProtection: Option<()>,
    bookViews: BookViews,
    sheets: Sheets,
    definedNames: Option<()>,
    calcPr: CalcPr,
}

#[derive(Debug, Deserialize)]
struct FileVersion {
    appName: String,
}

#[derive(Debug, Deserialize)]
struct WorkbookPr {
    showObjects: String,
    date1904: String,
}

serde_xlsx_items_struct!(BookViews, "workbookView" => WorkbookView);

#[derive(Debug, Deserialize)]
struct WorkbookView {
    showHorizontalScroll: String,
    showVerticalScroll: String,
    showSheetTabs: String,
    tabRatio: String,
    windowHeight: String,
    windowWidth: String,
    xWindow: String,
    yWindow: String,
}

serde_xlsx_items_struct!(Sheets, "sheet" => Sheet);

#[derive(Debug, Deserialize)]
struct Sheet {
    name: String,
    sheetId: String,
    
    //#[serde(rename = "r:id", default)]
    id: String,

    state: String,
}

#[derive(Debug, Deserialize)]
struct CalcPr {
    iterateCount: String,
    refMode: String,
    iterateDelta: String,
}

impl_from_xml_str!(Workbook);

//test_load_from_xml_str!(Workbook, "tests/xlsx/xl/workbook.xml");
