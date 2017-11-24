#![allow(non_snake_case, dead_code, unused_macros, non_upper_case_globals)]

/// serde_xlsx_items_struct!
///
/// ```rust
/// serde_xlsx_items_struct!(
///     NumFmts,
///     "numFmt" => NumFmt,
/// );
/// ```
/// generate:
/// ```rust
/// #[derive(Debug, Deserialize)]
/// struct NumFmts {
///     #[serde(rename = "numFmt", default)]
///     items: Vec<NumFmt>,
/// }
/// ```
///
/// ```rust
/// serde_xlsx_items_struct!(
///     NumFmts,
///     "numFmt" => NumFmt,
///     count: String
/// );
/// ```
/// generate:
/// ```rust
/// #[derive(Debug, Deserialize)]
/// struct NumFmts {
///     #[serde(rename = "numFmt", default)]
///     items: Vec<NumFmt>,
///     count: String
/// }
/// ```
///
/// ```rust
/// serde_xlsx_items_struct!{
///     name: SharedString,
///     item: "si" => StringItem,
///     fields: {
///         uniqueCount: String,
///         ...
///     }
/// };
/// ```
/// generate:
/// ```rust
/// #[derive(Debug, Deserialize)]
/// struct SharedString {
///     #[serde(rename = "si", default)]
///     items: Vec<StringItem>,
///     uniqueCount: String,
///     ...
/// }
/// ```
///
macro_rules! serde_xlsx_items_struct {
    ($struct_name:ident,
     $serde_name:tt  => $items_struct_name:ident
    ) => {
        serde_xlsx_items_struct!(
            name: $struct_name,
            item: $serde_name => $items_struct_name,
            fields: {}
        );
    };
    ($struct_name:ident,
     $serde_name:tt  => $items_struct_name:ident,
     $($element: ident: $ty: ty),*
    ) => {
        serde_xlsx_items_struct!(
            name: $struct_name,
            item: $serde_name => $items_struct_name,
            fields: {
                $($element: $ty),*
            }
        );
    };
    (name: $struct_name:ident,
     item: $serde_name:tt  =>  $items_struct_name:ident,
     fields: {
         $($element: ident: $ty: ty),*
     }) => {
        #[derive(Debug, Deserialize)]
        pub struct $struct_name {
            #[serde(rename = $serde_name, default)]
            pub items: Vec<$items_struct_name>,
            $(pub $element: $ty),*
        }

        /*impl $struct_name {
            pub fn items(self: &Self) -> &Vec<$items_struct_name> {
                return &self.items;
            }
        }*/
    };
}

macro_rules! impl_from_xml_str {
    ($struct_name:ident) => {
        const $struct_name: () = {
            use serde_xml_rs::{deserialize, Error};
            use std::io::{Read};

            impl $struct_name {
                pub fn from_xml_str(str: &str) -> Result<$struct_name, Error> {
                    deserialize(str.as_bytes())
                }

                pub fn from_xml<R: Read>(reader: R) -> Result<$struct_name, Error> {
                    deserialize(reader)
                }
            }
        };
    }
}

macro_rules! test_load_from_xml_str{
    ($struct_name:ident, $xml_file_path:tt) => {
        #[test]
        fn load_from_xml_str() {
            use std::io::prelude::*;
            use std::fs::File;

            let path = format!("{}/{}", env!("CARGO_MANIFEST_DIR"), $xml_file_path);
            match File::open(&path) {
                Ok(mut file) => {
                    let mut contents = String::new();
                    match file.read_to_string(&mut contents) {
                        Ok(_) => {
                            match $struct_name::from_xml_str(&contents) {
                                Ok(ss) => println!("{:#?}", ss),
                                Err(err) => println!("{:#?}", err),
                            }
                        }
                        Err(err) => println!("read file error: {}", err),
                    }
                }
                Err(err) => println!("open file error: {}", err),
            }
        }
    };
}

#[derive(Debug, Deserialize)]
pub struct Value {
    val: String,
}

pub mod styles;
pub mod shared_strings;
pub mod workbook;
pub mod theme;
pub mod rels;
pub mod sheet;
//mod content_types;
//mod doc_props;
