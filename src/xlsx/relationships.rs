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
    rename = "Relationships",
    prefix = "", 
    default_namespace = "",
    namespace = "http://schemas.openxmlformats.org/package/2006/relationships",
)]
pub struct RelationshipItems {
    #[yaserde(rename = "Relationship")]
    items: Vec<Relationship>
}

#[derive(Debug, YaDeserialize, YaSerialize)]
pub struct Relationship {
    #[yaserde(attribute, rename = "Id")]
    id: String,
    #[yaserde(attribute, rename = "Type")]
    typ: String,
    #[yaserde(attribute, rename = "Target")]
    target: String,
}

const TYPE_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

impl Relationships {
    pub fn find_by_id(&self, id: &str) -> Option<&Relationship> {
        self.relationships.items.iter().find(|rel| rel.id == id)
    }

    pub fn find_by_type(&self, typ: &str) -> Option<&Relationship> {
        self.relationships.items.iter()
            .find(|rel| rel.typ == format!("{}{}", TYPE_NS, typ))
    }

    pub fn add(&mut self, typ: &str, target: &str) {
        let id = format!("rId{}", self.relationships.items.len() + 1);
        let rel = Relationship {
            id,
            typ: TYPE_NS.to_string() + typ,
            target: target.to_string(),
        };
        self.relationships.items.push(rel);
    }
}

#[test]
fn test_load_ar() -> super::XlsxResult<()> {
    let mut ar = super::test::test_archive()?;

    println!("{}\n", Relationships::archive_str(&mut ar)?);

    let mut relationships = Relationships::load_archive(&mut ar)?;
    println!("{:?}\n", relationships.relationships);

    println!("{}\n", relationships.to_string()?);

    println!("{:?}\n", relationships.find_by_type("sharedStrings"));

    relationships.add("sharedStrings", "sharedStrings.xml");

    println!("{}\n", relationships.to_string()?);

    Ok(())
}