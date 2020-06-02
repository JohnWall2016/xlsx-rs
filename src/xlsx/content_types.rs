use super::{LoadArchive, XlsXResult, load_from_zip};
use super::zip::Archive;
use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

const NAME: &str = "[Content_Types].xml";

pub struct ContentTypes {
    types: Types
}

impl LoadArchive for ContentTypes {
    fn load_archive(ar: &mut Archive) -> XlsXResult<Self> {
        Ok(ContentTypes{ types: load_from_zip(ar, NAME)? })
    }
}

#[derive(Debug, YaDeserialize, YaSerialize)]
#[yaserde(prefix = "", namespace = "http://schemas.openxmlformats.org/package/2006/content-types")]
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

/*
#[derive(Debug, YaDeserialize)]
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
    None,
}
*/
/*
impl std::default::Default for Content {
    fn default() -> Content {
        Content::None
    }
}
*/

#[test]
fn test_load() -> super::XlsXResult<()> {
    let mut ar = Archive::new(super::test_file())?;

    use super::zip::ReadAll;
    let buf = ar.by_name(NAME)?.read_all_to_string()?;
    println!("{}", buf);

    let content_type = ContentTypes::load_archive(&mut ar)?;
    println!("{:?}", content_type.types);
/*
    for item in  &content_type.types.defaults {
        /*match item {
            Content::Default{extension, content_type} => {
                println!("Default: {} => {}", extension, content_type);
            },
            Content::Override{part_name, content_type} => {
                println!("Override: {} => {}", part_name, content_type);
            },
            Content::None => {},
        }*/
        println!("{:?}", item)
    }
*/

    use yaserde::ser::{to_string, to_string_content};
    println!("{}", to_string(&content_type.types)?);

    Ok(())
}