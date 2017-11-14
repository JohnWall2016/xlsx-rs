use serde_xml_rs::deserialize;

#[derive(Debug, Deserialize)]
struct StyleSheet {
    //name: String,
    numFmts: NumFmts,
    fonts: Fonts,
    fills: Fills
}

#[derive(Debug, Deserialize)]
struct NumFmts {
    count: String,
    
    #[serde(rename = "numFmt", default)]
    items: Vec<NumFmt>,
}

#[derive(Debug, Deserialize)]
struct NumFmt {
    numFmtId: String,
    formatCode: String,
}

#[derive(Debug, Deserialize)]
struct Fonts {
    count: String,
    
    #[serde(rename = "font", default)]
    items: Vec<Font>,
}

#[derive(Debug, Deserialize)]
struct Font {
    sz: Value,
    name: Value,
    family: Option<Value>,
    charset: Value,
    color: Option<Color>,
    b: Option<()>,
    u: Option<()>,
    i: Option<()>
}

#[derive(Debug, Deserialize)]
struct Value {
    #[serde(rename = "val", default)]
    value: String
}

#[derive(Debug, Deserialize)]
struct Color {
    rgb: Option<String>,
    indexed: Option<String>
}

#[derive(Debug, Deserialize)]
struct Fills {
    count: String,
    
    #[serde(rename = "fill", default)]
    items: Vec<Fill>,
}

#[derive(Debug, Deserialize)]
struct Fill {
    patternFill: PatternFill
}

#[derive(Debug, Deserialize)]
struct PatternFill {
    patternType: String, //attr
    fgColor: Option<Color>,
    bgColor: Option<Color>,
}


#[test]
fn load_xlsx_style() {
    use std::io::prelude::*;
    use std::fs::File;
    
    let path = format!("{}/tests/styles.xml", env!("CARGO_MANIFEST_DIR"));
    match File::open(&path) {
        Ok(mut file) => {
            let mut contents = String::new();
            match file.read_to_string(&mut contents) {
                Ok(_) => {
                    //println!("{}", contents);
                    let ss: StyleSheet = deserialize(contents.as_bytes()).unwrap();
                    println!("{:#?}", ss);
                }
                Err(err) => {
                    println!("read file error: {}", err)
                }
            }
        },
        Err(err) => {
            println!("open file error: {}", err)
        }
    }
}
