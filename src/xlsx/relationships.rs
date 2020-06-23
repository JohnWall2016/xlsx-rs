use super::{XlsxResult, ArchiveDeserable};
use crate::ar_deserable;

pub struct Relationships {
    relationships: internal::Relationships,
}

ar_deserable!(Relationships, "xl/_rels/workbook.xml.rels", relationships: internal::Relationships);

mod internal {
    use std::io::{Read, Write};
    use yaserde::{YaDeserialize, YaSerialize};

    #[derive(Debug, YaDeserialize, YaSerialize)]
    #[yaserde(
        prefix = "", 
        default_namespace = "",
        namespace = "http://schemas.openxmlformats.org/package/2006/relationships",
    )]
    pub struct Relationships {
        #[yaserde(rename = _)]
        pub contents: Vec<Relationship>
    }

    #[derive(Debug, YaDeserialize, YaSerialize)]
    pub struct Relationship {
        #[yaserde(attribute, rename = "Id")]
        pub id: String,

        #[yaserde(attribute, rename = "Type")]
        pub typ: String,

        #[yaserde(attribute, rename = "Target")]
        pub target: String,
    }
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