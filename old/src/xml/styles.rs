use super::Value;

#[derive(Debug, Deserialize)]
pub struct StyleSheet {
    //xmlns: String,
    pub numFmts: Option<NumFmts>,
    fonts: Fonts,
    fills: Fills,
    borders: Borders,
    cellStyleXfs: CellStyleXfs,
    pub cellXfs: CellXfs,
    cellStyles: CellStyles,
}


serde_xlsx_items_struct!(NumFmts, "numFmt" => NumFmt, count: String);

#[derive(Debug, Deserialize)]
pub struct NumFmt {
    pub numFmtId: String,
    pub formatCode: String,
}

serde_xlsx_items_struct!(Fonts, "font" => Font, count: String);

#[derive(Debug, Deserialize)]
pub struct Font {
    sz: Value,
    name: Value,
    family: Option<Value>,
    charset: Value,
    color: Option<Color>,
    b: Option<()>,
    u: Option<()>,
    i: Option<()>,
}

#[derive(Debug, Deserialize)]
struct Color {
    rgb: Option<String>,
    indexed: Option<String>,
}

serde_xlsx_items_struct!(Fills, "fill" => Fill, count: String);

#[derive(Debug, Deserialize)]
pub struct Fill {
    patternFill: PatternFill,
}

#[derive(Debug, Deserialize)]
struct PatternFill {
    patternType: String, //attr
    fgColor: Option<Color>,
    bgColor: Option<Color>,
}

serde_xlsx_items_struct!(Borders, "border" => Border, count: String);

#[derive(Debug, Deserialize)]
pub struct Border {
    left: Side,
    right: Side,
    top: Side,
    bottom: Side,
}

#[derive(Debug, Deserialize)]
struct Side {
    style: Option<String>,
    color: Option<Color>,
}

serde_xlsx_items_struct!(CellStyleXfs, "xf" => Xf, count: String);
serde_xlsx_items_struct!(CellXfs, "xf" => Xf, count: String);

#[derive(Debug, Deserialize)]
pub struct Xf {
    applyAlignment: String,
    applyBorder: String,
    applyFont: String,
    applyFill: String,
    applyNumberFormat: String,
    applyProtection: String,
    borderId: String,
    fillId: String,
    fontId: String,
    pub numFmtId: String,
    alignment: Alignment,
    xfId: Option<String>
}

#[derive(Debug, Deserialize)]
struct Alignment {
    horizontal: String,
    indent: String,
    shrinkToFit: String,
    textRotation: String,
    vertical: String,
    wrapText: String,
}

serde_xlsx_items_struct!(CellStyles, "cellStyle" => CellStyle, count: String);

#[derive(Debug, Deserialize)]
pub struct CellStyle {
    name: String,
    xfId: String,
}

impl_from_xml_str!(StyleSheet);

//test_load_from_xml_str!(StyleSheet, "test_data/xlsx/xl/styles.xml");
