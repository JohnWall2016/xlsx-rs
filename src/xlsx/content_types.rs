serde_xlsx_items_struct!(Types, "$value" => Type);

#[derive(Debug, Deserialize)]
enum Type {
    Override {
        #[serde(rename = "PartName")]
        partName: String,
        #[serde(rename = "ContentType")]
        contentType: String,
    },
    
    Default {
        #[serde(rename = "Extension")]
        partName: String,
        #[serde(rename = "ContentType")]
        contentType: String,
    },
}

impl_from_xml_str!(Types);

//test_load_from_xml_str!(Types, "tests/xlsx/[Content_Types].xml");
