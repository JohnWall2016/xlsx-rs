use super::{ArchiveDeserable, XlsxResult};
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

/*
pub struct AppProperties {
    properties: Properties,
}

impl ArchiveDeserable<Types> for AppProperties {
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

#[derive(Debug, YaDeserialize)]
struct Properties {
    #[serde(rename = "$value")]
    contents: Vec<Property>
}

#[derive(Debug, Deserialize)]
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
}

#[derive(Debug, Deserialize)]
struct Vector {
    size: String,
    #[serde(rename = "baseType")]
    base_type: String,
    #[serde(rename = "variant")]
    values: Vec<BaseType>,
}

#[derive(Debug, Deserialize)]
enum BaseType {
    #[serde(rename = "variant")]
    Variant(Variant),
    #[serde(rename = "lpstr")]
    Lpstr(String),
    #[serde(rename = "i4")]
    I4(String),
}

#[derive(Debug, Deserialize)]
enum Variant {
    #[serde(rename = "lpstr")]
    Lpstr(String),
    #[serde(rename = "i4")]
    I4(String),
}
*/

#[test]
fn test_load() -> XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    use super::zip::ReadAll;
    let buf = ar.by_name("docProps/app.xml")?.read_all_to_string()?;
    println!("{}", buf);

    Ok(())
}