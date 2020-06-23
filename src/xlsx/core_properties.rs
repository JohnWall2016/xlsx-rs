use super::{ArchiveDeserable, XlsxResult};
use crate::{ar_deserable, enum_default};
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

pub struct CoreProperties {
    properties: Properties,
}

ar_deserable!(CoreProperties, "docProps/core.xml", properties: Properties);

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "cp",
    default_namespace = "",
    namespace = "cp: http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    namespace = "dc: http://purl.org/dc/elements/1.1/",
    namespace = "dcterms: http://purl.org/dc/terms/",
    namespace = "dcmitype: http://purl.org/dc/dcmitype/",
    namespace = "xsi: http://www.w3.org/2001/XMLSchema-instance",
)]
struct Properties {
    #[yaserde(rename = _)]
    contents: Vec<Property>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
enum Property {
    #[yaserde(
        prefix = "dc",
        namespace = "dc: http://purl.org/dc/elements/1.1/",
        rename = "title"
    )]
    Title(String),

    #[yaserde(
        prefix = "dc",
        namespace = "dc: http://purl.org/dc/elements/1.1/",
        rename = "subject"
    )]
    Subject(String),

    #[yaserde(
        prefix = "dc",
        namespace = "dc: http://purl.org/dc/elements/1.1/",
        rename = "creator"
    )]
    Creator(String),

    #[yaserde(
        prefix = "dc",
        namespace = "dc: http://purl.org/dc/elements/1.1/",
        rename = "description"
    )]
    Description(String),

    #[yaserde(
        prefix = "cp",
        namespace = "cp: http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        rename = "keywords"
    )]
    Keywords(String),

    #[yaserde(
        prefix = "cp",
        namespace = "cp: http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        rename = "category"
    )]
    Category(String),

    None,
}

enum_default!(Property, None);

#[test]
fn test_load() -> XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", CoreProperties::archive_str(&mut ar)?);

    let core_properties = CoreProperties::load_archive(&mut ar)?;
    println!("{:?}\n", core_properties.properties);

    println!("{}\n", core_properties.to_string()?);

    Ok(())
}