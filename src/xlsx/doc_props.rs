#[derive(Debug, Deserialize)]
struct App { // Properties
    #[serde(rename = "TotalTime")]
    totalTime: String,

    #[serde(rename = "Application")]
    application: String,
}

impl_from_xml_str!(App);

//test_load_from_xml_str!(App, "tests/xlsx/docProps/app.xml");

#[derive(Debug, Deserialize)]
struct Core { // cp:coreProperties
}

impl_from_xml_str!(Core);

//test_load_from_xml_str!(Core, "tests/xlsx/docProps/core.xml");
