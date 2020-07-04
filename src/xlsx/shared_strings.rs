use std::io::{Read, Write};
use std::collections::HashMap;
use yaserde::{YaDeserialize, YaSerialize};
use crate::enum_default;
use super::{XlsxResult, ArchiveDeserable, YaDeserable};

use crate::xlsx::zip;
use yaserde::de::from_reader;

pub struct SharedStrings {
    strings: SharedStringItems,
    indexes: HashMap<String, usize>,
}

//ar_deserable!(SharedStrings, "xl/sharedStrings.xml", strings: SharedStringItems);
impl ArchiveDeserable for SharedStrings {
    fn path() -> &'static str {
        "xl/sharedStrings.xml"
    }

    fn from_archive(ar: &mut zip::Archive) -> XlsxResult<SharedStrings> {
        let strings: SharedStringItems = from_reader(ar.by_name(Self::path())?)?;

        let mut indexes = HashMap::new();
        for (index, item) in (&strings.items).iter().enumerate() {
            match item {
                SharedStringItem::Text(s) => {
                    indexes.insert(s.clone(), index);
                },
                _ => {}
            }
        }

        Ok(SharedStrings{
            strings,
            indexes,
        })
    }

    fn to_string(&self) -> XlsxResult<String> {
        Ok(self.strings.to_string()?)
    }
}

impl SharedStrings {
    pub fn get_string_by_index(&self, index: usize) -> Option<String> {
        use std::fmt::Write;
        if let Some(item) = self.strings.items.get(index) {
            return match item {
                SharedStringItem::Text(s) => Some(s.clone()),
                SharedStringItem::RichText(v) => {
                    let mut s = String::new();
                    for r in v {
                        write!(s, "{}", r.text).unwrap();
                    }
                    Some(s)
                },
                SharedStringItem::None => None
            }
        }
        None
    }

    pub fn get_index_for_string(&mut self, s: &str) -> usize {
        if let Some(index) = self.indexes.get(s) {
            *index
        } else {
            let index = self.strings.items.len();
            self.strings.count += 1;
            self.strings.unique_count += 1;
            self.strings.items.push(SharedStringItem::Text(s.to_string()));
            self.indexes.insert(s.to_string(), index);
            index
        }
    }
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    rename = "sst",
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)]
pub struct SharedStringItems {
    #[yaserde(attribute)]
    count: usize,

    #[yaserde(attribute, rename = "uniqueCount")]
    unique_count: usize,

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

    println!("{}\n", SharedStrings::archive_string(&mut ar)?);

    let shared_strings = SharedStrings::from_archive(&mut ar)?;
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

    use super::YaDeserable;

    let strings = SharedStringItems::from_str(s)?;
    println!("{:?}\n", strings);

    println!("{}\n", strings.to_string()?);

    Ok(())
}