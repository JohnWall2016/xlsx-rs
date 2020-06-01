use super::zip::Archive;
use super::{LoadArchive, Result, load_from_zip};

const NAME: &str = "docProps/app.xml";

pub struct AppProperties {
    properties: Properties,
}

impl LoadArchive for AppProperties {
    fn load_archive(ar: &mut Archive) -> Result<Self> {
        Ok(AppProperties{ properties: load_from_zip(ar, NAME)? })
    }
}

#[derive(Debug, Deserialize)]
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

#[test]
fn test_load() -> Result<()> {
    let mut ar = Archive::new(super::test_file())?;

    use super::zip::ReadAll;
    let buf = ar.by_name(NAME)?.read_all_to_string()?;
    println!("{}", buf);

    let app_properties = AppProperties::load_archive(&mut ar)?;
    println!("{:?}", app_properties.properties);

    Ok(())
}