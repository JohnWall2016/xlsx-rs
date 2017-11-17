#[derive(Debug, Deserialize)]
struct Worksheet {
    sheetPr: SheetPr,
    dimension: Dimension,
    sheetViews: SheetViews,
    cols: Cols,
    sheetData: SheetData,
    printOptions: PrintOptions,
    pageMargins: PageMargins,
    pageSetup: PageSetup,
    headerFooter: HeaderFooter,
}

#[derive(Debug, Deserialize)]
struct SheetPr {
    filterMode: String,
    pageSetUpPr: PageSetUpPr,
}

#[derive(Debug, Deserialize)]
struct PageSetUpPr {
    fitToPage: String,
}

#[derive(Debug, Deserialize)]
struct Dimension {
    #[serde(rename = "ref")]
    refer: String,
}

serde_xlsx_items_struct!(SheetViews, "sheetView" => SheetView);

#[derive(Debug, Deserialize)]
struct SheetView {
    windowProtection: String,
    showFormulas: String,
    showGridLines: String,
    showRowColHeaders: String,
    showZeros: String,
    rightToLeft: String,
    tabSelected: String,
    showOutlineSymbols: String,
    view: String,
    topLeftCell: String,
    zoomScale: String,
    zoomScaleNormal: String,
    zoomScalePageLayoutView: String,
    workbookViewId: String,

    selection: Selection,
}

#[derive(Debug, Deserialize)]
struct Selection {
    pane: String,
    activeCell: String,
    activeCellId: String,
    sqref: String,
}

serde_xlsx_items_struct!(Cols, "col" => ColDef);

#[derive(Debug, Deserialize)]
struct ColDef {
    collapsed: String,
    hidden: String,
    max: String,
    min: String,
    style: String,
    width: String,
    customWidth: String,
}

serde_xlsx_items_struct!(SheetData, "row" => Row);

serde_xlsx_items_struct!(
    Row, "c" => Col,
    r: String, ht: String,
    customHeight: String
);

#[derive(Debug, Deserialize)]
struct Col {
    r: String,
    s: String,
    t: Option<String>,

    v: Option<String>,
}

#[derive(Debug, Deserialize)]
struct PrintOptions {
    horizontalCentered: String,
    verticalCentered: String,
}

#[derive(Debug, Deserialize)]
struct PageMargins {
    left: String,
    right: String,
    top: String,
    bottom: String,
    header: String,
    footer: String,
}

#[derive(Debug, Deserialize)]
struct PageSetup {
    paperSize: String,
    orientation: String,
    horizontalDpi: String,
    verticalDpi: String,
}

#[derive(Debug, Deserialize)]
struct HeaderFooter {
    differentFirst: String,
    differentOddEven: String,
}

impl_from_xml_str!(Worksheet);

test_load_from_xml_str!(Worksheet, "tests/xlsx/xl/worksheets/sheet1.xml");
