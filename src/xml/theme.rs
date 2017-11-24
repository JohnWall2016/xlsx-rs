use super::Value;

// https://msdn.microsoft.com/zh-cn/library/office/documentformat.openxml.drawing.aspx

#[derive(Debug, Deserialize)]
pub struct Theme {
    pub themeElements: ThemeElements,
    objectDefaults: ObjectDefaults,
    extraClrSchemeLst: Option<()>,
}

#[derive(Debug, Deserialize)]
pub struct ThemeElements {
    pub clrScheme: ClrScheme,
    fontScheme: FontScheme,
    fmtScheme: FmtScheme,
}

#[derive(Debug, Deserialize)]
pub struct ClrScheme {
    name: String,
    
    pub accent1: Clr,
    pub accent2: Clr,
    pub accent3: Clr,
    pub accent4: Clr,
    pub accent5: Clr,
    pub accent6: Clr,
    pub dk1: Clr,
    pub dk2: Clr,
    pub folHlink: Clr,
    pub hlink: Clr,
    pub lt1: Clr,
    pub lt2: Clr,
}

impl IntoIterator for ClrScheme {
    type Item = (&'static str, Clr);
    type IntoIter = ::std::vec::IntoIter<Self::Item>;

    fn into_iter(self) -> Self::IntoIter {
        vec![
            ("accent1", self.accent1),
            ("accent2", self.accent2),
            ("accent3", self.accent3),
            ("accent4", self.accent4),
            ("accent5", self.accent5),
            ("accent6", self.accent6),
            ("dk1", self.dk1),
            ("dk2", self.dk2),
            ("folHlink", self.folHlink),
            ("hlink", self.hlink),
            ("lt1", self.lt1),
            ("lt2", self.lt2),
        ].into_iter()
    }
}

#[derive(Debug, Deserialize)]
pub enum Clr {
    #[serde(rename = "sysClr")]
    SysClr { val: String, lastClr: String },
    #[serde(rename = "srgbClr")]
    SrgbClr { val: String, alpha: Option<Value> },
    /*#[serde(rename = "schemeClr")]
    SchemeClr {
        val: String,
        tint: Option<Value>,
        satMod: Option<Value>,
        shade: Option<Value>,
    }*/
}

impl Clr {
    pub fn rgb_color(self: &Self) -> &str {
        match *self {
            Clr::SysClr { val: _, lastClr: ref clr } => clr,
            Clr::SrgbClr { val: ref clr, alpha: _ } => clr,
        }
    }
}

#[derive(Debug, Deserialize)]
pub struct SrgbClr { val: String, alpha: Option<Value> }

#[derive(Debug, Deserialize)]
pub struct SchemeClr {
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
pub struct FontType {
    typeface: String,
}

#[derive(Debug, Deserialize)]
pub struct Font {
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
pub enum FillStyle {
    #[serde(rename = "solidFill")]
    SolidFill { schemeClr: SchemeClr },
    #[serde(rename = "gradFill")]
    GradFill { rotWithShape: String, gsLst: GsLst, lin: Option<Lin>, path: Option<Path> },
}

serde_xlsx_items_struct!(GsLst, "gs" => Gs);

#[derive(Debug, Deserialize)]
pub struct Gs {
    pos: String,
    schemeClr: SchemeClr,
}

#[derive(Debug, Deserialize)]
pub struct Lin {
    ang: String,
    scaled: String,
}

#[derive(Debug, Deserialize)]
pub struct Path {
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
pub struct Ln {
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
pub struct EffectStyle {
    effectLst: EffectLst,
    scene3d: Option<Scene3d>,
    sp3d: Option<Sp3d>,
}

serde_xlsx_items_struct!(EffectLst, "$value" => Effect);

#[derive(Debug, Deserialize)]
pub enum Effect {
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

//test_load_from_xml_str!(Theme, "tests/xlsx/xl/theme/theme1.xml");
