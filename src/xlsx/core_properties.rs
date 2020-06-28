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
    rename = "coreProperties",
    prefix = "cp",
    default_namespace = "",
    namespace = "cp: http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    namespace = "dc: http://purl.org/dc/elements/1.1/",
    namespace = "dcterms: http://purl.org/dc/terms/",
    namespace = "dcmitype: http://purl.org/dc/dcmitype/",
    namespace = "xsi: http://www.w3.org/2001/XMLSchema-instance",
)]
pub struct Properties {
    #[yaserde(rename = _)]
    items: Vec<Property>,
}

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
enum Property {
    #[yaserde(prefix = "dc", rename = "title")]
    Title(String),

    #[yaserde(prefix = "dc", rename = "subject")]
    Subject(String),

    #[yaserde(prefix = "dc", rename = "creator")]
    Creator(String),

    #[yaserde(prefix = "dc", rename = "description")]
    Description(String),

    #[yaserde(prefix = "cp", rename = "keywords")]
    Keywords(String),

    #[yaserde(prefix = "cp",rename = "category")]
    Category(String),

    None,
}

enum_default!(Property::None);

#[test]
fn test_load() -> XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", CoreProperties::archive_string(&mut ar)?);

    let core_properties = CoreProperties::from_archive(&mut ar)?;
    println!("{:?}\n", core_properties.properties);

    println!("{}\n", core_properties.to_string()?);

    Ok(())
}

#[test]
fn test_load_str() -> XlsxResult<()> {
    let xml = r#"
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>Title</dc:title>
    <dc:subject>Subject</dc:subject>
    <dc:creator>Creator</dc:creator>
    <cp:keywords>Keywords</cp:keywords>
    <dc:description>Description</dc:description>
    <cp:category>Category</cp:category>
    </cp:coreProperties>
    "#;

    use super::YaDeserable;

    let properties  = Properties::from_str(xml)?;
    println!("{:?}\n", properties);

    println!("{}\n", properties.to_string()?);

    Ok(())
}