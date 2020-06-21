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
    #[yaserde(rename = _)]
    contents: Vec<Content>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "", 
    default_namespace = "", 
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
)]
enum Content {
    Default {
        #[yaserde(attribute, rename="Extension")]
        extension: String,
        #[yaserde(attribute, rename="ContentType")]
        content_type: String,
    },
    Override {
        #[yaserde(attribute, rename="PartName")]
        part_name: String,
        #[yaserde(attribute, rename="ContentType")]
        content_type: String,
    },
    Test(String),
    None,
}

impl std::default::Default for Content {
    fn default() -> Self {
        Self::None
    }
}

#[test]
fn test_load_ar() -> super::XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", ContentTypes::archive_str(&mut ar)?);

    let content_type = ContentTypes::load_archive(&mut ar)?;
    println!("{:?}\n", content_type.types);

    println!("{}\n", content_type.to_string()?);

    Ok(())
}

#[test]
fn test_load_str() -> super::XlsxResult<()> {
    let s = r#"
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>
<!--contents>
<Test>test string2</Test>
</contents-->
<Test>test string2</Test>
<Default Extension="rels1" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Default>
<Default Extension="rels2" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Default>
<Default Extension="rels3" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Default>
<Default Extension="rels4" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Default>
</Types>
    "#;
    let content_type = ContentTypes::load_string(s)?;
    println!("{:?}\n", content_type.types);

    println!("{}\n", content_type.to_string()?);

    Ok(())
}