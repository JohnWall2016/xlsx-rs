
serde_xlsx_items_struct!(Relationships, "Relationship" => Relationship);

#[derive(Debug, Deserialize)]
pub struct Relationship {
    #[serde(rename = "Id")]
    pub id: String,

    #[serde(rename = "Target")]
    pub target: String,

    #[serde(rename = "Type")]
    typ: String,
}

impl_from_xml_str!(Relationships);

//test_load_from_xml_str!(Relationships, "test_data/xlsx/xl/_rels/workbook.xml.rels");
