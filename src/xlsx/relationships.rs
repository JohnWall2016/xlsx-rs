use super::{XlsxResult, ArchiveDeserable};
use crate::ar_deserable;
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

pub struct Relationships {
    relationships: RelationshipItems,
}

ar_deserable!(Relationships, "xl/_rels/workbook.xml.rels", relationships: RelationshipItems);

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(
    prefix = "", 
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/package/2006/relationships",
)]
struct RelationshipItems {
    #[yaserde(rename = "Relationship")]
    items: Vec<Relationship>
}

#[derive(Debug, YaDeserialize, YaSerialize)]
struct Relationship {
    #[yaserde(attribute, rename = "Id")]
    id: String,
    #[yaserde(attribute, rename = "Type")]
    typ: String,
    #[yaserde(attribute, rename = "Target")]
    target: String,
}

#[test]
fn test_load_ar() -> super::XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", Relationships::archive_str(&mut ar)?);

    let relationships = Relationships::load_archive(&mut ar)?;
    println!("{:?}\n", relationships.relationships);

    println!("{}\n", relationships.to_string()?);

    Ok(())
}