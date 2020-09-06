use super::base::{ArchiveDeserable, XlsxResult};
use crate::{ar_deserable, enum_default};
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

pub struct ContentTypes {
    types: Types,
}

ar_deserable!(ContentTypes, "[Content_Types].xml", types: Types);

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "",
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
)]
pub struct Types {
    #[yaserde(rename = _)]
    items: Vec<Type>,
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "",
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
)]
pub enum Type {
    Default {
        #[yaserde(attribute, rename = "Extension")]
        extension: String,
        #[yaserde(attribute, rename = "ContentType")]
        content_type: String,
    },
    Override {
        #[yaserde(attribute, rename = "PartName")]
        part_name: String,
        #[yaserde(attribute, rename = "ContentType")]
        content_type: String,
    },
    Test(String),
    None,
}

enum_default!(Type::None);

impl ContentTypes {
    pub fn find_by_part_name(&self, part_name: &str) -> Option<&Type> {
        self.types.items.iter().find(|ty| {
            if let Type::Override { part_name: pn, .. } = ty {
                if pn == part_name {
                    return true;
                }
            }
            false
        })
    }

    pub fn add(&mut self, part_name: &str, content_type: &str) {
        let ty = Type::Override {
            part_name: part_name.to_string(),
            content_type: content_type.to_string(),
        };
        self.types.items.push(ty);
    }
}

#[test]
fn test_load_ar() -> XlsxResult<()> {
    let mut ar = super::base::test::test_archive()?;

    println!("{}\n", ContentTypes::archive_string(&mut ar)?);

    let mut content_type = ContentTypes::from_archive(&mut ar)?;
    println!("{:?}\n", content_type.types);

    println!("{}\n", content_type.to_string()?);

    println!(
        "{:?}\n",
        content_type.find_by_part_name("/xl/sharedStrings.xml")
    );

    content_type.add(
        "/xl/sharedStrings.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
    );

    println!("{}\n", content_type.to_string()?);

    Ok(())
}

#[test]
fn test_load_str() -> XlsxResult<()> {
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
    
    use super::base::YaDeserable;

    let types = Types::from_str(s)?;
    println!("{:?}\n", types);

    println!("{}\n", types.to_string()?);

    Ok(())
}
