serde_xlsx_items_struct!{
    name: SharedStrings,
    item: "si" => StringItem,
    fields: {
        count: String,
        uniqueCount: String
    }
}

#[derive(Debug, Deserialize)]
pub struct StringItem {
    pub t: String,
}

impl_from_xml_str!(SharedStrings);

//test_load_from_xml_str!(SharedString, "tests/xlsx/xl/sharedStrings.xml");
