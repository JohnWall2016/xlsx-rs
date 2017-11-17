
serde_xlsx_items_struct!(Relationships, "Relationship" => Relationship);

#[derive(Debug, Deserialize)]
struct Relationship {
    #[serde(rename = "Id")]
    id: String,

    #[serde(rename = "Target")]
    target: String,

    #[serde(rename = "Type")]
    typ: String,
}

impl_from_xml_str!(Relationships);

test_load_from_xml_str!(Relationships, "tests/xlsx/xl/_rels/workbook.xml.rels");
