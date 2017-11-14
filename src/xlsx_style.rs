#[derive(Debug, Deserialize)]
struct StyleSheet {
    //xmlns: String,
    numFmts: NumFmts,
    fonts: Fonts,
    fills: Fills,
    borders: Borders,
    cellStyleXfs: CellStyleXfs,
    cellXfs: CellXfs,
    cellStyles: CellStyles,
}

/// serde_count_items_struct!
/// 
/// ```rust
/// serde_count_items_struct!(NumFmts, "numFmt", NumFmt);
/// ```
/// generate:
/// ```rust
/// #[derive(Debug, Deserialize)]
/// struct NumFmts {
///     count: String,
///
///     #[serde(rename = "numFmt", default)]
///     items: Vec<NumFmt>,
/// }
/// ``` 
///
macro_rules! serde_count_items_struct {
    ($struct_name:ident, $serde_name:tt, $items_struct_name:ident) => {
        #[derive(Debug, Deserialize)]
        struct $struct_name {
            count: String,

            #[serde(rename = $serde_name, default)]
            items: Vec<$items_struct_name>,
        }
    }
}

serde_count_items_struct!(NumFmts, "numFmt", NumFmt);

#[derive(Debug, Deserialize)]
struct NumFmt {
    numFmtId: String,
    formatCode: String,
}

serde_count_items_struct!(Fonts, "font", Font);

#[derive(Debug, Deserialize)]
struct Font {
    sz: Value,
    name: Value,
    family: Option<Value>,
    charset: Value,
    color: Option<Color>,
    b: Option<()>,
    u: Option<()>,
    i: Option<()>,
}

#[derive(Debug, Deserialize)]
struct Value {
    #[serde(rename = "val", default)]
    value: String,
}

#[derive(Debug, Deserialize)]
struct Color {
    rgb: Option<String>,
    indexed: Option<String>,
}

serde_count_items_struct!(Fills, "fill", Fill);

#[derive(Debug, Deserialize)]
struct Fill {
    patternFill: PatternFill,
}

#[derive(Debug, Deserialize)]
struct PatternFill {
    patternType: String, //attr
    fgColor: Option<Color>,
    bgColor: Option<Color>,
}

serde_count_items_struct!(Borders, "border", Border);

#[derive(Debug, Deserialize)]
struct Border {
    left: Side,
    right: Side,
    top: Side,
    bottom: Side,
}

#[derive(Debug, Deserialize)]
struct Side {
    style: Option<String>,
    color: Option<Color>,
}

serde_count_items_struct!(CellStyleXfs, "xf", Xf);
serde_count_items_struct!(CellXfs, "xf", Xf);

#[derive(Debug, Deserialize)]
struct Xf {
    applyAlignment: String,
    applyBorder: String,
    applyFont: String,
    applyFill: String,
    applyNumberFormat: String,
    applyProtection: String,
    borderId: String,
    fillId: String,
    fontId: String,
    numFmtId: String,
    alignment: Alignment,
    xfId: Option<String>
}

#[derive(Debug, Deserialize)]
struct Alignment {
    horizontal: String,
    indent: String,
    shrinkToFit: String,
    textRotation: String,
    vertical: String,
    wrapText: String,
}

serde_count_items_struct!(CellStyles, "cellStyle", CellStyle);

#[derive(Debug, Deserialize)]
struct CellStyle {
    name: String,
    xfId: String,
}

use serde_xml_rs::{deserialize, Error};

impl StyleSheet {
    fn from_xml_str(str: &String) -> Result<Self, Error> {
        deserialize(str.as_bytes())
    }
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
                    //let ss: StyleSheet = deserialize(contents.as_bytes()).unwrap();
                    match StyleSheet::from_xml_str(&contents) {
                        Ok(ss) => println!("{:#?}", ss),
                        Err(err) => println!("{:#?}", err)
                    }
                    
                }
                Err(err) => println!("read file error: {}", err),
            }
        }
        Err(err) => println!("open file error: {}", err),
    }
}
