use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};
use crate::{ar_deserable, enum_default};
use super::{XlsxResult, ArchiveDeserable};

pub struct SharedStrings {
    strings: SharedStringItems,
}

ar_deserable!(SharedStrings, "xl/sharedStrings.xml", strings: SharedStringItems);

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "sst",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
struct SharedStringItems {
    #[yaserde(attribute)]
    count: u32,

    #[yaserde(attribute, rename = "uniqueCount")]
    unique_count: u32,

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
    RichText(Vec<RichTextRun>),

    None,
}

enum_default!(SharedStringItem::None);

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "r",
    prefix = "",
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
struct RichTextRun {
    #[yaserde(rename = "rPr")]
    properties: Option<RunProperies>,

    #[yaserde(rename = "t")]
    text: String,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
struct RunProperies {
    #[yaserde(rename = _)]
    items: Vec<RunProperty>,
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
    FontSize { #[yaserde(attribute)] val: String },

    #[yaserde(rename = "color")]
    Color { #[yaserde(attribute)] theme: String },

    #[yaserde(rename = "rFont")]
    RunFont { #[yaserde(attribute)] val: String },

    #[yaserde(rename = "family")]
    FontFamily { #[yaserde(attribute)] val: String },

    #[yaserde(rename = "scheme")]
    Scheme { #[yaserde(attribute)] val: String },

    Unknown,
}

enum_default!(RunProperty::Unknown);

#[test]
fn test_load_ar() -> super::XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", SharedStrings::archive_str(&mut ar)?);

    let shared_strings = SharedStrings::load_archive(&mut ar)?;
    println!("{:?}\n", shared_strings.strings);

    println!("{}\n", shared_strings.to_string()?);

    Ok(())
}

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
    println!("{:?}\n", shared_strings.strings);

    println!("{}\n", shared_strings.to_string()?);

    Ok(())
}