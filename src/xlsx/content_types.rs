use super::{LoadArchive, Result, load_from_zip};
use super::zip::Archive;

const NAME: &str = "[Content_Types].xml";

pub struct ContentTypes {
    types: Types
}

impl LoadArchive for ContentTypes {
    fn load_archive(ar: &mut Archive) -> Result<Self> {
        Ok(ContentTypes{ types: load_from_zip(ar, NAME)? })
    }
}

#[derive(Debug, Deserialize)]
struct Types {
    #[serde(rename = "$value")]
    contents: Vec<Content>
}

#[derive(Debug, Deserialize)]
enum Content {
    Default {
        #[serde(rename="Extension")]
        extension: String,
        #[serde(rename="ContentType")]
        content_type: String,
    },
    Override {
        #[serde(rename="PartName")]
        part_name: String,
        #[serde(rename="ContentType")]
        content_type: String,
    }
}

#[test]
fn test_load() -> super::Result<()> {
    let mut ar = Archive::new(super::test_file())?;

    use super::zip::ReadAll;
    let buf = ar.by_name(NAME)?.read_all_to_string()?;
    println!("{}", buf);

    let content_type = ContentTypes::load_archive(&mut ar)?;
    println!("{:?}", content_type.types);

    for item in  &content_type.types.contents {
        match item {
            Content::Default{extension, content_type} => {
                println!("Default: {} => {}", extension, content_type);
            },
            Content::Override{part_name, content_type} => {
                println!("Override: {} => {}", part_name, content_type);
            }
        }
    }
    
    Ok(())
}