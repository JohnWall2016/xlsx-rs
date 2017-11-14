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
}

impl_from_xml_str!(Theme);

test_load_from_xml_str!(Theme, "tests/xlsx/xl/theme/theme1.xml");
