use super::{ArchiveDeserable, XlsxResult};
use crate::enum_default;
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

pub struct AppProperties {
    properties: Properties,
}

impl ArchiveDeserable<Properties> for AppProperties {
    fn path() -> &'static str {
        "docProps/app.xml"
    }

    fn deseralize_to(de: Properties) -> XlsxResult<Self> {
        Ok(AppProperties{ properties: de })
    }

    fn seralize_to(&self) -> XlsxResult<&Properties> {
        Ok(&self.properties)
    }
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "",
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    namespace = "vt: http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
)]
struct Properties {
    #[yaserde(rename = _)]
    contents: Vec<Property>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "", 
    namespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    namespace = "vt: http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
)]
enum Property {
    Application(String),
    DocSecurity(String),
    ScaleCrop(String),
    HeadingPairs {
        vector: Vector
    },
    TitlesOfParts {
        vector: Vector
    },
    Company,
    LinksUpToDate(String),
    SharedDoc(String),
    HyperlinksChanged(String),
    AppVersion(String),
    TotalTime(String),

    None,
}

enum_default!(Property, None);

#[derive(Debug, YaDeserialize, YaSerialize, Default)]
struct Vector {
    size: String,

    #[yaserde(rename = "baseType")]
    base_type: String,

    #[yaserde(rename = "variant")]
    values: Vec<BaseType>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
enum BaseType {
    #[yaserde(rename = "variant")]
    Variant(Variant),

    #[yaserde(rename = "lpstr")]
    Lpstr(String),

    #[yaserde(rename = "i4")]
    I4(String),

    None,
}

enum_default!(BaseType, None);

#[derive(Debug, YaDeserialize, YaSerialize)]
enum Variant {
    #[yaserde(rename = "lpstr")]
    Lpstr(String),

    #[yaserde(rename = "i4")]
    I4(String),

    None,
}

enum_default!(Variant, None);


#[test]
fn test_load() -> XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", AppProperties::archive_str(&mut ar)?);

    let app_properties = AppProperties::load_archive(&mut ar)?;
    println!("{:?}\n", app_properties.properties);

    println!("{}\n", app_properties.to_string()?);

    Ok(())
}