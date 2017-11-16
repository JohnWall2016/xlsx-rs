use super::Value;

#[derive(Debug, Deserialize)]
struct Theme {
    themeElements: ThemeElements,
    extraClrSchemeLst: Option<()>,
}

#[derive(Debug, Deserialize)]
struct ThemeElements {
    clrScheme: ClrScheme,
    fontScheme: FontScheme,
    fmtScheme: FmtScheme,
}

// https://msdn.microsoft.com/zh-cn/library/office/documentformat.openxml.drawing.colorscheme.aspx
#[derive(Debug, Deserialize)]
struct ClrScheme {
    name: String,
    
    accent1: Clr,
    accent2: Clr,
    accent3: Clr,
    accent4: Clr,
    accent5: Clr,
    accent6: Clr,
    dk1: Clr,
    dk2: Clr,
    folHlink: Clr,
    hlink: Clr,
    lt1: Clr,
    lt2: Clr,
}

#[derive(Debug, Deserialize)]
enum Clr {
    #[serde(rename = "sysClr")]
    SysClr { val: String, lastClr: String },
    #[serde(rename = "srgbClr")]
    SrgbClr { val: String },
}

#[derive(Debug, Deserialize)]
pub struct FontScheme {
    name: String,

    majorFont: Fonts,
    minorFont: Fonts,
}

serde_xlsx_items_struct!(
    Fonts, "font" => Font,
    latin: FontType, ea: FontType, cs: FontType
);

#[derive(Debug, Deserialize)]
struct FontType {
    typeface: String,
}

#[derive(Debug, Deserialize)]
struct Font {
    script: String,
    typeface: String,
}

#[derive(Debug, Deserialize)]
struct FmtScheme {
    fillStyleLst: FillStyleLst,
    bgFillStyleLst: FillStyleLst,
    // TODO: lnStyleLst, effectStyleLst
}

#[derive(Debug, Deserialize)]
struct SchemeClr {
    val: String,
    tint: Option<Value>,
    satMod: Option<Value>,
    shade: Option<Value>,
}

serde_xlsx_items_struct!(FillStyleLst, "$value" => FillStyle);

#[derive(Debug, Deserialize)]
enum FillStyle {
    #[serde(rename = "solidFill")]
    SolidFill { schemeClr: SchemeClr },
    #[serde(rename = "gradFill")]
    GradFill { rotWithShape: String, gsLst: GsLst, lin: Option<Lin>, path: Option<Path> },
}

serde_xlsx_items_struct!(GsLst, "gs" => Gs);

#[derive(Debug, Deserialize)]
struct Gs {
    pos: String,
    schemeClr: SchemeClr,
}

#[derive(Debug, Deserialize)]
struct Lin {
    ang: String,
    scaled: String,
}

#[derive(Debug, Deserialize)]
struct Path {
    path: String,
    fillToRect: FillToRect,
}

#[derive(Debug, Deserialize)]
struct FillToRect {
    l: String,
    t: String,
    r: String,
    b: String,
}

impl_from_xml_str!(Theme);

test_load_from_xml_str!(Theme, "tests/xlsx/xl/theme/theme1.xml");
