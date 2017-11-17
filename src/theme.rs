use super::Value;

// https://msdn.microsoft.com/zh-cn/library/office/documentformat.openxml.drawing.aspx

#[derive(Debug, Deserialize)]
struct Theme {
    themeElements: ThemeElements,
    objectDefaults: ObjectDefaults,
    extraClrSchemeLst: Option<()>,
}

#[derive(Debug, Deserialize)]
struct ThemeElements {
    clrScheme: ClrScheme,
    fontScheme: FontScheme,
    fmtScheme: FmtScheme,
}

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
    SrgbClr { val: String, alpha: Option<Value> },
    #[serde(rename = "schemeClr")]
    SchemeClr {
        val: String,
        tint: Option<Value>,
        satMod: Option<Value>,
        shade: Option<Value>,
    }
}

#[derive(Debug, Deserialize)]
struct SrgbClr { val: String, alpha: Option<Value> }

#[derive(Debug, Deserialize)]
struct SchemeClr {
    val: String,
    tint: Option<Value>,
    satMod: Option<Value>,
    shade: Option<Value>,
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
    lnStyleLst: LnStyleLst,
    effectStyleLst: EffectStyleLst,
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

serde_xlsx_items_struct!(LnStyleLst, "$value" => Ln);

#[derive(Debug, Deserialize)]
struct Ln {
    w: String,
    cap: String,
    cmpd: String,
    algn: String,

    solidFill: SolidFill,
    prstDash: Value,
}

#[derive(Debug, Deserialize)]
struct SolidFill { schemeClr: SchemeClr }

serde_xlsx_items_struct!(EffectStyleLst, "$value" => EffectStyle);

#[derive(Debug, Deserialize)]
struct EffectStyle {
    effectLst: EffectLst,
    scene3d: Option<Scene3d>,
    sp3d: Option<Sp3d>,
}

serde_xlsx_items_struct!(EffectLst, "$value" => Effect);

#[derive(Debug, Deserialize)]
enum Effect {
    #[serde(rename = "outerShdw")]
    OuterShdw {
        blurRad: String,
        dist: String,
        dir: String,
        rotWithShape: String,

        srgbClr: SrgbClr,
    }
}

#[derive(Debug, Deserialize)]
struct Scene3d {
    camera: Camera,
    lightRig: LightRig,
}

#[derive(Debug, Deserialize)]
struct Camera {
    prst: String,
    rot: Rot,
}

#[derive(Debug, Deserialize)]
struct Rot {
    lat: String,
    lon: String,
    rev: String,
}

#[derive(Debug, Deserialize)]
struct LightRig {
    rig: String,
    dir: String,
    rot: Rot,
}

#[derive(Debug, Deserialize)]
struct Sp3d {
    bevelT: BevelT,
}

#[derive(Debug, Deserialize)]
struct BevelT {
    w: String,
    h: String,
}

#[derive(Debug, Deserialize)]
struct ObjectDefaults {
    spDef: Def,
    lnDef: Def,
}

#[derive(Debug, Deserialize)]
struct Def {
    spPr: Option<()>,
    bodyPr: Option<()>,
    lstStyle: Option<()>,

    style: Style,
}

#[derive(Debug, Deserialize)]
struct Style {
    lnRef: Ref,
    fillRef: Ref,
    effectRef: Ref,
    fontRef: Ref,
}

#[derive(Debug, Deserialize)]
struct Ref {
    idx: String,
    schemeClr: SchemeClr,
}

impl_from_xml_str!(Theme);

test_load_from_xml_str!(Theme, "tests/xlsx/xl/theme/theme1.xml");
