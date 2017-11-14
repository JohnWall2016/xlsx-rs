serde_xlsx_items_struct!{
    name: SharedString,
    item: "si" => StringItem,
    fields: {
        count: String,
        uniqueCount: String
    }
}

#[derive(Debug, Deserialize)]
struct StringItem {
    t: String,
}

impl_from_xml_str!(SharedString);

//test_load_from_xml_str!(SharedString, "tests/xlsx/xl/sharedStrings.xml");
