use super::{XlsxResult, ArchiveDeserable};
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

pub struct ContentTypes {
    types: Types
}

impl ArchiveDeserable<Types> for ContentTypes {
    fn path() -> &'static str {
        "[Content_Types].xml"
    }

    fn deseralize_to(de: Types) -> XlsxResult<Self> {
        Ok(ContentTypes{ types: de })
    }

    fn seralize_to(&self) -> XlsxResult<&Types> {
        Ok(&self.types)
    }
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
)]
struct Types {
    #[yaserde(rename = "Override")]
    overrides: Vec<Override>,
    #[yaserde(rename = "Default")]
    defaults: Vec<Default>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
struct Default {
    #[yaserde(attribute, rename="Extension")]
    extension: String,
    #[yaserde(attribute, rename="ContentType")]
    content_type: String,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
struct Override {
    #[yaserde(attribute, rename="PartName")]
    part_name: String,
    #[yaserde(attribute, rename="ContentType")]
    content_type: String,
}

#[test]
fn test_load() -> super::XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}", ContentTypes::archive_str(&mut ar)?);

    let content_type = ContentTypes::load_archive(&mut ar)?;
    println!("{:?}", content_type.types);

    println!("{}", content_type.to_string()?);

    Ok(())
}