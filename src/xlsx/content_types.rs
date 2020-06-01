pub struct ContentTypes {
    types: Types
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
    let mut ar = super::zip::Archive::new(super::test_file())?;
    let types: Types = super::load_from_zip(&mut ar, "[Content_Types].xml")?;
    println!("{:?}", types);
    Ok(())
}