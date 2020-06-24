use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};
use crate::{ar_deserable, enum_default};
use super::{XlsxResult, ArchiveDeserable};

pub struct SharedStrings {
    shared_strings: SharedStringItems,
}

ar_deserable!(SharedStrings, "xl/sharedStrings.xml", shared_strings: SharedStringItems);

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "sst",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
struct SharedStringItems {
    #[yaserde(attribute)]
    count: String,

    #[yaserde(attribute, rename = "uniqueCount")]
    unique_count: String,

    #[yaserde(rename = "si")]
    items: Vec<SharedStringItem>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "si",
    prefix = "",
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
enum SharedStringItem {
    #[yaserde(rename = "t")]
    Text(String),

    #[yaserde(rename = "r")]
    RichTextRuns(Vec<RichTextRuns>),

    None,
}

enum_default!(SharedStringItem::None);

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    rename = "r",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
struct RichTextRuns {
    //#[yaserde(rename = _)]
    //properties: Vec<RichTextRun>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "r",
    prefix = "",
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
enum RichTextRun {
    #[yaserde(rename = "rPr")]
    RunProperties(RunProperties),

    #[yaserde(rename = "t")]
    Text(String),

    None,
}

enum_default!(RichTextRun::None);

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
#[yaserde(
    rename = "rPr",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
struct RunProperties {
    #[yaserde(rename = _)]
    properties: Vec<RunProperty>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
enum RunProperty {
    #[yaserde(rename = "b")]
    Bold,

    #[yaserde(rename = "sz")]
    FontSize { val: String },

    #[yaserde(rename = "color")]
    Color { theme: String },

    #[yaserde(rename = "rFont")]
    RunFont { val: String },

    #[yaserde(rename = "family")]
    FontFamily { val: String },

    #[yaserde(rename = "scheme")]
    Scheme { val: String },

    None,
}

enum_default!(RunProperty::None);

#[test]
fn test_load_str() -> super::XlsxResult<()> {
    let s = r#"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="13" uniqueCount="4">
	<si>
		<t>Foo</t>
	</si>
	<si>
		<t>Bar</t>
	</si>
	<si>
		<t>Goo</t>
	</si>
	<si>
		<r>
			<t>s</t>
		</r>
        <r>
			<rPr>
				<b/>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr>
            <t>d;</t>
		</r>
        <r>
            <rPr>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr>
            <t>lfk;l</t>
		</r>
	</si>
</sst>
"#;

    let shared_strings = SharedStrings::load_string(s)?;
    println!("{:?}\n", shared_strings.shared_strings);

    println!("{}\n", shared_strings.to_string()?);

    Ok(())
}