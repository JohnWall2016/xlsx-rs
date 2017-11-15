use std::collections::BTreeMap as Map;

#[derive(Debug, Deserialize)]
struct Theme {
    themeElements: ThemeElements,
    extraClrSchemeLst: Option<()>,
}

#[derive(Debug, Deserialize)]
struct ThemeElements {
    clrScheme: ClrScheme,
}

#[derive(Debug, Deserialize)]
struct ClrScheme {
    name: String,
    //#[serde(rename = "$value", default)]
    //clrs: Map<String, Clr>,
    dk1: Clr,
}

#[derive(Debug, Deserialize)]
enum Clr {
    #[serde(rename = "sysClr")]
    SysClr{val: String, lastClr: String},
    #[serde(rename = "srgbClr")]
    SrgbClr{val: String},
}

impl_from_xml_str!(Theme);

test_load_from_xml_str!(Theme, "tests/xlsx/xl/theme/theme1.xml");
