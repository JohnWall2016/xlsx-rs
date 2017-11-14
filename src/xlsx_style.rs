use serde_xml_rs::deserialize;

#[derive(Debug, Deserialize)]
struct StyleSheet {
    //pub xmlns: String,
    pub numFmts: NumFmts,
    pub fonts: Fonts,
    pub fills: Fills
}

#[derive(Debug, Deserialize)]
struct NumFmts {
    pub count: String,
    
    #[serde(rename = "numFmt", default)]
    pub items: Vec<NumFmt>,
}

#[derive(Debug, Deserialize)]
struct NumFmt {
    pub numFmtId: String,
    pub formatCode: String,
}

#[derive(Debug, Deserialize)]
struct Fonts {
    pub count: String,
    
    #[serde(rename = "font", default)]
    pub items: Vec<Font>,
}

#[derive(Debug, Deserialize)]
struct Font {
    pub sz: Value,
    pub name: Value,
    pub family: Option<Value>,
    pub charset: Value,
    pub color: Option<Color>,
    pub b: Option<()>,
    pub u: Option<()>,
    pub i: Option<()>
}

#[derive(Debug, Deserialize)]
struct Value {
    #[serde(rename = "val", default)]
    pub value: String
}

#[derive(Debug, Deserialize)]
enum Color {
    pub rgb: String,
    pub index: String
}

#[derive(Debug, Deserialize)]
struct Fills {
    pub count: String,
    
    #[serde(rename = "fill", default)]
    pub items: Vec<Fill>,
}

#[derive(Debug, Deserialize)]
struct Fill {
    pub patternFill: PatternFill
}

#[derive(Debug, Deserialize)]
struct PatternFill {
    pub patternType: String, //attr
    pub fgColor: Color,
    pub bgColor: 
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
